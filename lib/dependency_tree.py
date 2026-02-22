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
    RevitLinkInstance,
    ExternalFileReferenceType,
    ModelPathUtils,
)

from cloud_helpers import build_cloud_model_path, make_detached_open_options


# ------------------------------------------------------------------
# Read direct link info from an already-open document
# ------------------------------------------------------------------

def _is_empty_guid(guid_str):
    """Return True if the GUID string is all zeros."""
    return guid_str.replace("-", "").replace("0", "") == ""


def get_direct_link_guids(doc):
    """Return a list of dicts describing each *direct* cloud link in *doc*.

    Uses RevitLinkInstance.GetLinkDocument() to read cloud GUIDs from the
    in-memory linked document itself.  This is more reliable than
    RevitLinkType.GetExternalFileReference() for ACC models, which can
    return empty GUIDs when the link reference is unresolved.

    Falls back to ExternalFileReference for links whose document is not
    loaded in memory.

    Each dict contains:
        name, project_guid, model_guid, user_path

    Also returns a list of diagnostic strings for links that were skipped.
    """
    results = []
    skipped = []
    seen_guids = set()

    # --- Phase 1: collect link info from RevitLinkInstance (preferred) ---
    instances = list(
        FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    )
    instance_type_ids = set()  # track which type IDs we already handled

    for inst in instances:
        inst_name = "<unknown>"
        try:
            inst_name = inst.Name
        except Exception:
            pass

        # Skip nested links
        try:
            lt = doc.GetElement(inst.GetTypeId())
            if lt is not None and lt.IsNestedLink:
                skipped.append("Skipped (nested): {0}".format(inst_name))
                continue
        except Exception:
            pass

        # Avoid processing same link type twice
        try:
            tid = inst.GetTypeId()
            if tid in instance_type_ids:
                continue
            instance_type_ids.add(tid)
        except Exception:
            pass

        linked_doc = None
        try:
            linked_doc = inst.GetLinkDocument()
        except Exception as ex:
            skipped.append("GetLinkDocument failed: {0} -- {1}".format(
                inst_name, ex))

        if linked_doc is not None:
            try:
                cloud_mp = linked_doc.GetCloudModelPath()
                pg = str(cloud_mp.GetProjectGUID())
                mg = str(cloud_mp.GetModelGUID())

                if _is_empty_guid(mg):
                    skipped.append("Skipped (empty model GUID from doc): {0}".format(
                        inst_name))
                    continue

                if mg in seen_guids:
                    continue
                seen_guids.add(mg)

                up = ""
                try:
                    up = ModelPathUtils.ConvertModelPathToUserVisiblePath(cloud_mp)
                except Exception:
                    pass

                name = linked_doc.Title or inst_name
                results.append({
                    "name": name,
                    "project_guid": pg,
                    "model_guid": mg,
                    "user_path": up,
                })
                continue  # success — skip fallback
            except Exception as ex:
                skipped.append(
                    "GetCloudModelPath failed: {0} -- {1}".format(inst_name, ex))

        # --- Fallback: try ExternalFileReference on the link type ---
        try:
            lt = doc.GetElement(inst.GetTypeId())
            if lt is None:
                continue
            efr = lt.GetExternalFileReference()
            if efr is None:
                continue
            if efr.ExternalFileReferenceType != ExternalFileReferenceType.RevitLink:
                continue
            model_path = efr.GetAbsolutePath()
            if model_path is None:
                continue
            pg = str(model_path.GetProjectGUID())
            mg = str(model_path.GetModelGUID())
            if _is_empty_guid(mg):
                skipped.append("Skipped (empty GUID from EFR): {0}".format(inst_name))
                continue
            if mg in seen_guids:
                continue
            seen_guids.add(mg)
            up = ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
            results.append({
                "name": lt.Name or inst_name,
                "project_guid": pg,
                "model_guid": mg,
                "user_path": up,
            })
        except Exception as ex:
            skipped.append("EFR fallback failed: {0} -- {1}".format(inst_name, ex))

    # --- Phase 2: catch link types with no instances ---
    link_types = list(
        FilteredElementCollector(doc).OfClass(RevitLinkType).ToElements()
    )
    for lt in link_types:
        lt_name = "<unknown>"
        try:
            lt_name = lt.Name
        except Exception:
            pass
        try:
            if lt.IsNestedLink:
                continue
        except Exception:
            pass
        if lt.Id in instance_type_ids:
            continue  # already handled above
        try:
            efr = lt.GetExternalFileReference()
            if efr is None:
                continue
            if efr.ExternalFileReferenceType != ExternalFileReferenceType.RevitLink:
                continue
            model_path = efr.GetAbsolutePath()
            if model_path is None:
                continue
            pg = str(model_path.GetProjectGUID())
            mg = str(model_path.GetModelGUID())
            if _is_empty_guid(mg) or mg in seen_guids:
                continue
            seen_guids.add(mg)
            up = ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
            results.append({
                "name": lt_name,
                "project_guid": pg,
                "model_guid": mg,
                "user_path": up,
            })
        except Exception:
            pass

    if not results and not instances and not link_types:
        skipped.append("No RevitLinkType or RevitLinkInstance elements found.")

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
# Discover children from in-memory linked documents
# ------------------------------------------------------------------

def _get_loaded_link_docs(doc):
    """Return a dict mapping model_guid → linked Document for every
    RevitLinkInstance whose document is currently loaded in memory."""
    link_docs = {}
    for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements():
        try:
            linked_doc = inst.GetLinkDocument()
            if linked_doc is None:
                continue
            # Extract cloud GUIDs from the linked document
            cloud_mp = linked_doc.GetCloudModelPath()
            mg = str(cloud_mp.GetModelGUID())
            if mg not in link_docs:
                link_docs[mg] = linked_doc
        except Exception:
            continue
    return link_docs


def build_dependency_tree_from_doc(root_doc, root_region, root_project_guid,
                                   root_model_guid, root_name="ROOT",
                                   progress_callback=None):
    """Build the full dependency tree from the already-open root document
    by walking in-memory linked documents (no server round-trips).

    This avoids opening child models from scratch, which can fail on ACC
    when the central server cannot be reached from a cold open.
    """
    adjacency = defaultdict(list)
    model_info = {}
    visited = set()
    queue = deque()

    # Collect ALL loaded link documents from the root
    all_link_docs = _get_loaded_link_docs(root_doc)

    if progress_callback:
        progress_callback("Found {0} loaded link document(s) in memory.".format(
            len(all_link_docs)))

    queue.append((root_region, root_project_guid, root_model_guid,
                  root_name, root_doc))
    visited.add(root_model_guid)

    while queue:
        region, proj, mod, name, doc = queue.popleft()
        model_info[mod] = {
            "name": name,
            "project_guid": proj,
            "region": region,
        }

        if progress_callback:
            progress_callback("Scanning: {0}".format(name))

        children = []
        if doc is not None:
            try:
                children, skipped = get_direct_link_guids(doc)

                # Also collect loaded link docs from this child document
                if doc != root_doc:
                    child_link_docs = _get_loaded_link_docs(doc)
                    for k, v in child_link_docs.items():
                        if k not in all_link_docs:
                            all_link_docs[k] = v

                if skipped and progress_callback:
                    for s in skipped:
                        progress_callback("  ⚠ " + s)

                if progress_callback:
                    progress_callback(
                        "  Found {0} cloud link(s) in {1}".format(
                            len(children), name))

            except Exception as ex:
                model_info[mod]["error"] = str(ex)
                if progress_callback:
                    progress_callback(
                        "  ✗ Error scanning {0}: {1}".format(name, ex))
        else:
            # Document not loaded — we know it exists but cannot scan its
            # sub-links.  Record it without an error so it still appears in
            # the sync list (it just won't have children).
            if progress_callback:
                progress_callback(
                    "  ⚠ {0}: not loaded in memory (sub-links unknown)".format(name))

        for child in children:
            child_key = child["model_guid"]
            adjacency[mod].append(child_key)

            if child_key not in visited:
                visited.add(child_key)
                # Look up the child's in-memory document (may be None)
                child_doc = all_link_docs.get(child_key)
                queue.append((
                    region,
                    child["project_guid"],
                    child["model_guid"],
                    child["name"],
                    child_doc,
                ))

    return dict(adjacency), dict(model_info)


# ------------------------------------------------------------------
# BFS tree builder (legacy — opens each model from scratch)
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
