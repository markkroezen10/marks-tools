# -*- coding: utf-8 -*-
"""Discover the full dependency tree of ACC cloud models and produce a
bottom-up (leaf-first) sync order via topological sort.

Usage
-----
    from dependency_tree import build_dependency_tree, sync_order

    adj, info = build_dependency_tree(app, region, proj_guid, model_guid)
    ordered   = sync_order(adj, info)
    # ordered is a list of model_guid strings, leaves first, root last.
"""

from collections import defaultdict, deque

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkType,
    ExternalFileReferenceType,
    ModelPathUtils,
)

from cloud_helpers import build_cloud_model_path, make_detached_open_options


# ------------------------------------------------------------------
# Read direct link GUIDs from an already-open document
# ------------------------------------------------------------------

def get_direct_link_guids(doc):
    """Return a list of dicts describing each *direct* RevitLinkType in *doc*.

    Nested links (``IsNestedLink == True``) are skipped so we only get the
    immediate children of this document.

    Each dict contains:
        name, project_guid, model_guid, user_path

    Also returns a list of diagnostic strings for links that were skipped.
    """
    results = []
    skipped = []
    link_types = list(
        FilteredElementCollector(doc).OfClass(RevitLinkType).ToElements()
    )
    if not link_types:
        skipped.append("No RevitLinkType elements found in document.")
        return results, skipped

    for lt in link_types:
        lt_name = "<unknown>"
        try:
            lt_name = lt.Name
        except Exception:
            pass

        # Skip nested (transitive) links
        try:
            if lt.IsNestedLink:
                skipped.append("Skipped (nested): {0}".format(lt_name))
                continue
        except Exception:
            pass

        try:
            efr = lt.GetExternalFileReference()
        except Exception as ex:
            skipped.append("Skipped (GetExternalFileReference failed): {0} — {1}".format(lt_name, ex))
            continue
        if efr is None:
            skipped.append("Skipped (no ExternalFileReference): {0}".format(lt_name))
            continue
        if efr.ExternalFileReferenceType != ExternalFileReferenceType.RevitLink:
            skipped.append("Skipped (not a RevitLink type): {0}".format(lt_name))
            continue

        model_path = efr.GetAbsolutePath()
        if model_path is None:
            skipped.append("Skipped (null absolute path): {0}".format(lt_name))
            continue

        # Extract cloud GUIDs — only succeeds for cloud model paths;
        # file-based paths will throw and be skipped.
        try:
            pg = str(model_path.GetProjectGUID())
            mg = str(model_path.GetModelGUID())
            up = ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
        except Exception as ex:
            skipped.append("Skipped (not a cloud path): {0} — {1}".format(lt_name, ex))
            continue

        results.append({
            "name": lt_name,
            "project_guid": pg,
            "model_guid": mg,
            "user_path": up,
        })
    return results, skipped


# ------------------------------------------------------------------
# Discover children of a single cloud model (open detached → scan → close)
# ------------------------------------------------------------------

def discover_children(app, region, project_guid, model_guid):
    """Open cloud model **detached** (fast), read direct links, close.

    Returns (children_list, skipped_list).
    """
    mp = build_cloud_model_path(region, project_guid, model_guid)
    opts = make_detached_open_options()
    doc = app.OpenDocumentFile(mp, opts)
    try:
        return get_direct_link_guids(doc)
    finally:
        try:
            doc.Close(False)
        except Exception:
            pass


# ------------------------------------------------------------------
# BFS tree builder
# ------------------------------------------------------------------

def build_dependency_tree(app, root_region, root_project_guid, root_model_guid,
                          root_name="ROOT", progress_callback=None,
                          root_doc=None):
    """Walk the cloud-link tree via BFS.

    Parameters
    ----------
    app : Autodesk.Revit.ApplicationServices.Application
        ``uiapp.Application`` or ``revit.app``.
    root_region, root_project_guid, root_model_guid : str
        Identifiers for the top-level model.
    root_name : str
        Display name for the root node.
    progress_callback : callable(msg) or None
        Called with a status string whenever a model is opened / closed.
    root_doc : Document or None
        If the root model is already open (e.g. active document), pass it
        here so the BFS reads links directly instead of trying to re-open it.

    Returns
    -------
    adjacency : dict
        ``{ model_guid: [child_model_guid, ...], ... }``
    model_info : dict
        ``{ model_guid: {"name", "project_guid", "region"}, ... }``
    """
    adjacency = defaultdict(list)
    model_info = {}
    visited = set()
    queue = deque()

    queue.append((root_region, root_project_guid, root_model_guid, root_name))
    visited.add(root_model_guid)

    while queue:
        region, proj, mod, name = queue.popleft()
        model_info[mod] = {
            "name": name,
            "project_guid": proj,
            "region": region,
        }

        if progress_callback:
            progress_callback("Scanning: {0}".format(name))

        children = []
        try:
            # Use the already-open document for the root model;
            # open-detached for every other model in the tree.
            if root_doc and mod == root_model_guid:
                children, skipped = get_direct_link_guids(root_doc)
            else:
                children, skipped = discover_children(app, region, proj, mod)

            # Report skipped links via progress callback
            if skipped and progress_callback:
                for s in skipped:
                    progress_callback("  ⚠ " + s)

            if progress_callback:
                progress_callback(
                    "  Found {0} cloud link(s) in {1}".format(len(children), name)
                )

        except Exception as ex:
            # Could not open this model — record it but continue BFS
            model_info[mod]["error"] = str(ex)
            if progress_callback:
                progress_callback(
                    "  ✗ Error scanning {0}: {1}".format(name, ex)
                )

        for child in children:
            child_key = child["model_guid"]
            adjacency[mod].append(child_key)

            if child_key not in visited:
                visited.add(child_key)
                # Assume same region unless cross-region linking is encountered
                queue.append((
                    region,
                    child["project_guid"],
                    child["model_guid"],
                    child["name"],
                ))

    return dict(adjacency), dict(model_info)


# ------------------------------------------------------------------
# Topological sort — leaves first, root last
# ------------------------------------------------------------------

def sync_order(adjacency, model_info):
    """DFS post-order traversal → children appear before parents.

    Returns a list of model_guid strings in the order they should be synced
    (leaf models first, root model last).
    """
    visited = set()
    order = []

    def _dfs(node):
        if node in visited:
            return
        visited.add(node)
        for child in adjacency.get(node, []):
            _dfs(child)
        order.append(node)

    # Start from every node to handle disconnected components
    for node in model_info:
        _dfs(node)

    return order


# ------------------------------------------------------------------
# Pretty-print tree (for debug / UI)
# ------------------------------------------------------------------

def format_tree(adjacency, model_info, root_guid, indent=0):
    """Return a list of (indent_level, model_guid, display_name) tuples."""
    lines = []
    info = model_info.get(root_guid, {})
    name = info.get("name", root_guid[:8] + "...")
    lines.append((indent, root_guid, name))
    for child_guid in adjacency.get(root_guid, []):
        lines.extend(format_tree(adjacency, model_info, child_guid, indent + 1))
    return lines
