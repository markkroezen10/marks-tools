# -*- coding: utf-8 -*-
# pyRevit - Get Cloud Project GUID + Model GUID immediately from the ACTIVE document (no closing)
# Uses Document.GetCloudModelPath if available; otherwise uses .NET reflection to find GUID accessors.

from pyrevit import revit, forms
import System
from System import Guid

doc = revit.doc
if not doc:
    forms.alert("No active document.", exitscript=True)

def copy_clipboard(txt):
    try:
        System.Windows.Clipboard.SetText(txt)
        return True
    except Exception:
        return False

def safe_str(x):
    try:
        return str(x)
    except Exception:
        return "<unprintable>"

# Get user-visible path for info
user_path = ""
try:
    from Autodesk.Revit.DB import ModelPathUtils
    cp = None
    try:
        cp = doc.GetWorksharingCentralModelPath()
    except Exception:
        cp = None
    if cp:
        user_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(cp)
except Exception:
    pass

project_guid = None
model_guid = None
region = ""

# -----------------------------
# 1) Try Document.GetCloudModelPath()
# -----------------------------
cloudmp = None
try:
    if hasattr(doc, "GetCloudModelPath"):
        cloudmp = doc.GetCloudModelPath()
except Exception:
    cloudmp = None

def try_extract_from_cloudmodelpath(obj):
    """Try common method/property names on a CloudModelPath-like object."""
    pg = mg = rg = None

    # methods
    for name in ["GetProjectGuid", "GetProjectGUID", "ProjectGuid", "ProjectGUID"]:
        try:
            m = getattr(obj, name)
            pg = m() if callable(m) else m
            break
        except Exception:
            pass

    for name in ["GetModelGuid", "GetModelGUID", "ModelGuid", "ModelGUID"]:
        try:
            m = getattr(obj, name)
            mg = m() if callable(m) else m
            break
        except Exception:
            pass

    for name in ["GetRegion", "Region"]:
        try:
            m = getattr(obj, name)
            rg = m() if callable(m) else m
            break
        except Exception:
            pass

    return pg, mg, rg

if cloudmp:
    pg, mg, rg = try_extract_from_cloudmodelpath(cloudmp)
    if pg and mg:
        project_guid = pg
        model_guid = mg
        region = safe_str(rg) if rg else ""

# -----------------------------
# 2) Reflection fallback on Document (find anything returning Guid)
# -----------------------------
reflection_notes = []
if not project_guid or not model_guid:
    try:
        t = doc.GetType()
        methods = t.GetMethods()
        cloud_methods = [m for m in methods if "Cloud" in m.Name or "cloud" in m.Name]
        reflection_notes.append("Doc cloud-related methods found: " + ", ".join([m.Name for m in cloud_methods]) if cloud_methods else "No doc cloud-related methods found.")

        # Try calling any zero-arg method that returns an object with project/model guid accessors
        for m in cloud_methods:
            try:
                if m.GetParameters().Length == 0:
                    obj = m.Invoke(doc, None)
                    if obj:
                        pg, mg, rg = try_extract_from_cloudmodelpath(obj)
                        if pg and mg:
                            project_guid = pg
                            model_guid = mg
                            region = safe_str(rg) if rg else ""
                            reflection_notes.append("Used doc method: " + m.Name)
                            break
            except Exception:
                continue
    except Exception as ex:
        reflection_notes.append("Reflection on doc failed: " + safe_str(ex))

# -----------------------------
# 3) If still not found: show diagnostics so we can pick exact API
# -----------------------------
if not project_guid or not model_guid:
    msg = "Could not extract GUIDs from the active model.\n\n"
    if user_path:
        msg += "User-visible path:\n{0}\n\n".format(user_path)
    msg += "Diagnostics:\n- " + "\n- ".join(reflection_notes) if reflection_notes else "No diagnostics."
    msg += "\n\nNext step: send me this diagnostics text; I will map the right API call for your Revit build."
    forms.alert(msg, exitscript=True)

# Normalize to strings
pg_str = safe_str(project_guid)
mg_str = safe_str(model_guid)

csv_row = "{0},{1},{2},{3}".format(doc.Title, region, pg_str, mg_str)
copied = copy_clipboard(csv_row)

forms.alert(
    "Model: {0}\n"
    "Region: {1}\n"
    "Project GUID: {2}\n"
    "Model GUID: {3}\n\n"
    "{4}\n\n"
    "CSV (copied):\n{5}\n\n"
    "User-visible path:\n{6}"
    .format(doc.Title, region, pg_str, mg_str,
            ("Copied to clipboard." if copied else "Clipboard copy failed."),
            csv_row,
            user_path),
    warn_icon=False
)
