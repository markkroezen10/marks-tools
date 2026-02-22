# -*- coding: utf-8 -*-
"""Get cloud Project GUID + Model GUID from the active document."""
__title__ = "Get Model ID"
__doc__ = "Extracts Project GUID, Model GUID, and Region from the active " \
          "ACC / BIM360 cloud document and copies a CSV row to the clipboard."
__author__ = "Mark"

from pyrevit import revit, forms
import System

from guid_extractor import extract_cloud_ids, get_user_visible_path, safe_str


doc = revit.doc
if not doc:
    forms.alert("No active document.", exitscript=True)


def copy_clipboard(txt):
    try:
        System.Windows.Clipboard.SetText(txt)
        return True
    except Exception:
        return False


# Extract IDs
try:
    region, pg_str, mg_str = extract_cloud_ids(doc)
except RuntimeError as ex:
    forms.alert(str(ex), exitscript=True)

user_path = get_user_visible_path(doc)

csv_row = "{0},{1},{2},{3}".format(doc.Title, region, pg_str, mg_str)
copied = copy_clipboard(csv_row)

forms.alert(
    "Model: {0}\n"
    "Region: {1}\n"
    "Project GUID: {2}\n"
    "Model GUID: {3}\n\n"
    "{4}\n\n"
    "CSV (copied):\n{5}\n\n"
    "User-visible path:\n{6}".format(
        doc.Title, region, pg_str, mg_str,
        "Copied to clipboard." if copied else "Clipboard copy failed.",
        csv_row,
        user_path,
    ),
    warn_icon=False,
)
