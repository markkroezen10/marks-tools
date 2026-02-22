# -*- coding: utf-8 -*-
# pyRevit (Revit 2026) - For each ACC cloud model (in order):
# 1) Open cloud model
# 2) Open CLOSED USER worksets (verified)
# 3) Reload Revit links ONE BY ONE (with delay + robust error handling)
# 4) ReloadLatest
# 5) SyncWithCentral + Relinquish
# 6) Close

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    ModelPathUtils, OpenOptions, WorksetConfiguration, WorksetConfigurationOption,
    SynchronizeWithCentralOptions, TransactWithCentralOptions, RelinquishOptions,
    FilteredWorksetCollector, WorksetKind, Transaction,
    FilteredElementCollector, RevitLinkType
)
from System import Guid
import time

# -----------------------------
# MODELS (ORDER MATTERS) - exactly as provided
# -----------------------------
MODELS = [
    ("BMR-ARC-00-XX-M3-BM-9100", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "0c7f141c-4f90-48e3-9f84-ad62707b9549"),
    ("BMR-ARC-00-XX-M3-BM-9101", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "e6754ac5-caf0-4050-b2b8-fa92f2ffcb31"),
    ("BMR-ARC-00-XX-M3-BM-9102", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "2f8fef20-3807-46a6-9f6e-f15f11fec342"),
    ("BMR-ARC-00-XX-M3-BM-9103", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "f5a23ebd-0ccb-4ce1-bcfd-97f098a6353a"),
    ("BMR-ARC-00-XX-M3-BM-9104", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "8329af7e-8068-4a93-9c99-65beeb6317c5"),
    ("BMR-ARC-00-XX-M3-BM-9105", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "17d844e9-1d3d-433e-beff-7587ee8744a3"),
    ("BMR-ARC-00-XX-M3-BM-9107", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "d803eb73-ae01-408b-a8c5-a3e78cbe1d8a"),
    ("BMR-ARC-00-XX-M3-BM-9001", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "71548a33-c4b9-4e50-8b04-d389c422146d"),
    ("BMR-ARC-00-XX-M3-BM-9002", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "66d61310-0951-4fa1-8ed8-f60cd1f0d85b"),
    ("BMR-ARC-00-XX-M3-BM-9003", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "62fcea84-f0df-4978-b425-516e3b4695cb"),
    ("BMR-ARC-00-XX-M3-BM-9004", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "a2d6aaba-9aa6-402c-b4a5-569acb1b9003"),
    ("BMR-ARC-00-XX-M3-BM-9005", "EMEA", "50f7a824-a024-4dfa-8645-b3993dc8f535", "e8849df2-89ee-4214-b1dc-494ba5447079"),
]

# -----------------------------
# SETTINGS (stability)
# -----------------------------
OPEN_ALL_WORKSETS_ON_OPEN = True
OPEN_CLOSED_USER_WORKSETS_AFTER_OPEN = True

RELOAD_LINKS_ONE_BY_ONE = True
LINK_RELOAD_DELAY_SECONDS = 2.0   # increase to 3-5 if you still see instability

RELOAD_LATEST_BEFORE_SYNC = True
SYNC_DELAY_SECONDS = 1.0          # small settle time before SWC

CLOSE_AFTER_SYNC = True           # keep stable


# -----------------------------
# HELPERS
# -----------------------------
def build_cloud_model_path(region_str, project_guid_str, model_guid_str):
    return ModelPathUtils.ConvertCloudGUIDsToCloudPath(
        region_str, Guid(project_guid_str), Guid(model_guid_str)
    )

def make_open_options(open_all_worksets=True):
    open_opts = OpenOptions()
    ws_mode = WorksetConfigurationOption.OpenAllWorksets if open_all_worksets else WorksetConfigurationOption.CloseAllWorksets
    open_opts.SetOpenWorksetsConfiguration(WorksetConfiguration(ws_mode))
    return open_opts

def _fresh_user_worksets(doc):
    return list(FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets())

def open_all_closed_user_worksets(doc):
    if not doc.IsWorkshared:
        return (0, 0)

    ws_table = doc.GetWorksetTable()
    before = _fresh_user_worksets(doc)
    closed_before = [ws for ws in before if not ws.IsOpen]

    if not closed_before:
        return (0, 0)

    t = Transaction(doc, "Open closed USER worksets")
    t.Start()
    try:
        for ws in closed_before:
            try:
                ws_table.OpenWorkset(ws.Id)
            except:
                pass
        t.Commit()
    except:
        try:
            t.RollBack()
        except:
            pass

    after = _fresh_user_worksets(doc)
    still_closed = [ws for ws in after if not ws.IsOpen]
    return (len(closed_before), len(still_closed))

def reload_links_one_by_one(doc, delay_s=2.0):
    """Reload each RevitLinkType individually, with its own transaction and delay.
       Returns (ok_names, fail_list)."""
    link_types = list(FilteredElementCollector(doc).OfClass(RevitLinkType).ToElements())
    ok = []
    fail = []

    # Sort by name for deterministic order
    try:
        link_types.sort(key=lambda x: x.Name)
    except:
        pass

    for lt in link_types:
        name = "<unknown>"
        try:
            name = lt.Name
        except:
            pass

        # Each link reload in its own transaction = more stable
        t = Transaction(doc, "Reload link: {0}".format(name))
        t.Start()
        try:
            lt.Reload()
            t.Commit()
            ok.append(name)
        except Exception as ex:
            try:
                t.RollBack()
            except:
                pass
            fail.append((name, str(ex)))

        # Give Revit a moment to process cloud/link activity
        if delay_s and delay_s > 0:
            time.sleep(delay_s)

    return (ok, fail)

def sync_with_central(doc):
    tcc_opts = TransactWithCentralOptions()
    sync_opts = SynchronizeWithCentralOptions()

    # Guard options (differs per build)
    try:
        sync_opts.SaveLocalBefore = True
    except:
        pass
    try:
        sync_opts.SaveLocalAfter = True
    except:
        pass

    rel = RelinquishOptions(True)
    rel.UserWorksets = True
    rel.ViewWorksets = True
    rel.FamilyWorksets = True
    rel.StandardWorksets = True
    rel.CheckedOutElements = True
    sync_opts.SetRelinquishOptions(rel)

    doc.SynchronizeWithCentral(tcc_opts, sync_opts)

def safe_close(doc):
    try:
        doc.Close(False)
    except:
        pass


# -----------------------------
# MAIN
# -----------------------------
confirm = forms.alert(
    "Per model:\n"
    "1) Open cloud model\n"
    "2) Open closed USER worksets\n"
    "3) Reload Revit links (one by one)\n"
    "4) ReloadLatest\n"
    "5) SyncWithCentral\n"
    "6) Close\n\n"
    "Proceed?",
    ok=False, yes=True, no=True
)
if not confirm:
    raise SystemExit

uiapp = revit.uidoc.Application
open_opts = make_open_options(OPEN_ALL_WORKSETS_ON_OPEN)

ok_models = []
fail_models = []

for (model_name, region, proj, mid) in MODELS:
    d = None
    try:
        mp = build_cloud_model_path(region, proj, mid)
        d = uiapp.Application.OpenDocumentFile(mp, open_opts)

        opened_cnt, still_closed_cnt = (0, 0)
        if OPEN_CLOSED_USER_WORKSETS_AFTER_OPEN:
            opened_cnt, still_closed_cnt = open_all_closed_user_worksets(d)

        link_ok = []
        link_fail = []
        if RELOAD_LINKS_ONE_BY_ONE:
            link_ok, link_fail = reload_links_one_by_one(d, delay_s=LINK_RELOAD_DELAY_SECONDS)

        if RELOAD_LATEST_BEFORE_SYNC:
            try:
                d.ReloadLatest()
            except:
                pass

        if SYNC_DELAY_SECONDS and SYNC_DELAY_SECONDS > 0:
            time.sleep(SYNC_DELAY_SECONDS)

        sync_with_central(d)

        ok_models.append(
            "{0} | worksets opened: {1} (still closed: {2}) | links reloaded: {3} (fails: {4})".format(
                model_name, opened_cnt, still_closed_cnt, len(link_ok), len(link_fail)
            )
        )

    except Exception as ex:
        fail_models.append((model_name, str(ex)))

    finally:
        if CLOSE_AFTER_SYNC and d:
            safe_close(d)

out = "DONE.\n\nSUCCESS:\n"
out += "\n".join(["- " + s for s in ok_models]) if ok_models else "- (none)"
out += "\n\nFAILED MODELS:\n"
out += "\n".join(["- {0}: {1}".format(n, e) for (n, e) in fail_models]) if fail_models else "- (none)"

forms.alert(out, warn_icon=False)
