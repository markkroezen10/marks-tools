# -*- coding: utf-8 -*-
"""Extract cloud GUIDs (project, model, region) from a Revit Document."""


def safe_str(x):
    try:
        return str(x)
    except Exception:
        return "<unprintable>"


def _try_extract_from_cloudmodelpath(obj):
    """Probe common property / method names on a CloudModelPath-like object.

    Returns (project_guid, model_guid, region) — any may be None.
    """
    pg = mg = rg = None

    for name in ("GetProjectGuid", "GetProjectGUID", "ProjectGuid", "ProjectGUID"):
        try:
            m = getattr(obj, name)
            pg = m() if callable(m) else m
            break
        except Exception:
            pass

    for name in ("GetModelGuid", "GetModelGUID", "ModelGuid", "ModelGUID"):
        try:
            m = getattr(obj, name)
            mg = m() if callable(m) else m
            break
        except Exception:
            pass

    for name in ("GetRegion", "Region"):
        try:
            m = getattr(obj, name)
            rg = m() if callable(m) else m
            break
        except Exception:
            pass

    return pg, mg, rg


def extract_cloud_ids(doc):
    """Return (region, project_guid_str, model_guid_str) for a cloud Document.

    Raises RuntimeError if GUIDs cannot be determined.
    """
    project_guid = None
    model_guid = None
    region = ""
    diagnostics = []

    # --- 1) GetCloudModelPath (preferred) ---
    cloudmp = None
    try:
        if hasattr(doc, "GetCloudModelPath"):
            cloudmp = doc.GetCloudModelPath()
    except Exception:
        cloudmp = None

    if cloudmp:
        pg, mg, rg = _try_extract_from_cloudmodelpath(cloudmp)
        if pg and mg:
            project_guid = pg
            model_guid = mg
            region = safe_str(rg) if rg else ""

    # --- 2) Reflection fallback ---
    if not project_guid or not model_guid:
        try:
            t = doc.GetType()
            methods = t.GetMethods()
            cloud_methods = [m for m in methods if "Cloud" in m.Name or "cloud" in m.Name]

            if cloud_methods:
                diagnostics.append(
                    "Cloud methods: " + ", ".join(m.Name for m in cloud_methods)
                )
            else:
                diagnostics.append("No cloud methods found on Document.")

            for m in cloud_methods:
                try:
                    if m.GetParameters().Length == 0:
                        obj = m.Invoke(doc, None)
                        if obj:
                            pg, mg, rg = _try_extract_from_cloudmodelpath(obj)
                            if pg and mg:
                                project_guid = pg
                                model_guid = mg
                                region = safe_str(rg) if rg else ""
                                diagnostics.append("Used method: " + m.Name)
                                break
                except Exception:
                    continue
        except Exception as ex:
            diagnostics.append("Reflection failed: " + safe_str(ex))

    if not project_guid or not model_guid:
        raise RuntimeError(
            "Could not extract cloud GUIDs.\nDiagnostics:\n- "
            + "\n- ".join(diagnostics) if diagnostics else "No diagnostics."
        )

    return (
        region,
        safe_str(project_guid),
        safe_str(model_guid),
    )


def get_user_visible_path(doc):
    """Best-effort user-visible path string for the document."""
    try:
        from Autodesk.Revit.DB import ModelPathUtils
        cp = None
        try:
            cp = doc.GetWorksharingCentralModelPath()
        except Exception:
            cp = None
        if cp:
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cp)
    except Exception:
        pass
    return ""
