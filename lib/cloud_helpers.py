# -*- coding: utf-8 -*-
"""Shared helpers for opening, syncing, and closing ACC/BIM360 cloud models."""

from Autodesk.Revit.DB import (
    ModelPathUtils,
    OpenOptions,
    WorksetConfiguration,
    WorksetConfigurationOption,
    DetachFromCentralOption,
    SynchronizeWithCentralOptions,
    TransactWithCentralOptions,
    RelinquishOptions,
    FilteredWorksetCollector,
    WorksetKind,
    FilteredElementCollector,
    RevitLinkType,
    Transaction,
)
from System import Guid
import time


# ------------------------------------------------------------------
# Cloud path helpers
# ------------------------------------------------------------------

def build_cloud_model_path(region, project_guid, model_guid):
    """Return a ModelPath for an ACC / BIM360 cloud model."""
    return ModelPathUtils.ConvertCloudGUIDsToCloudPath(
        region, Guid(project_guid), Guid(model_guid)
    )


# ------------------------------------------------------------------
# OpenOptions factories
# ------------------------------------------------------------------

def make_open_options(open_all_worksets=True):
    """Standard open: workshared, all (or no) worksets open."""
    opts = OpenOptions()
    ws_mode = (
        WorksetConfigurationOption.OpenAllWorksets
        if open_all_worksets
        else WorksetConfigurationOption.CloseAllWorksets
    )
    opts.SetOpenWorksetsConfiguration(WorksetConfiguration(ws_mode))
    return opts


def make_detached_open_options():
    """Fast read-only open (detached, all worksets closed).

    Use this during dependency discovery — it skips heavy data loading
    but still exposes RevitLinkType metadata.
    """
    opts = OpenOptions()
    opts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
    opts.SetOpenWorksetsConfiguration(
        WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
    )
    return opts


# ------------------------------------------------------------------
# Workset helpers
# ------------------------------------------------------------------

def _fresh_user_worksets(doc):
    return list(
        FilteredWorksetCollector(doc)
        .OfKind(WorksetKind.UserWorkset)
        .ToWorksets()
    )


def open_all_closed_user_worksets(doc):
    """Open every closed USER workset. Returns (closed_before, still_closed)."""
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
            except Exception:
                pass
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass

    after = _fresh_user_worksets(doc)
    still_closed = [ws for ws in after if not ws.IsOpen]
    return (len(closed_before), len(still_closed))


# ------------------------------------------------------------------
# Link reload
# ------------------------------------------------------------------

def reload_links_one_by_one(doc, delay_seconds=2.0):
    """Reload each RevitLinkType individually.

    Returns (ok_names, fail_list) where fail_list items are (name, error_msg).
    """
    link_types = list(
        FilteredElementCollector(doc).OfClass(RevitLinkType).ToElements()
    )
    try:
        link_types.sort(key=lambda x: x.Name)
    except Exception:
        pass

    ok = []
    fail = []

    for lt in link_types:
        name = "<unknown>"
        try:
            name = lt.Name
        except Exception:
            pass

        t = Transaction(doc, "Reload link: {0}".format(name))
        t.Start()
        try:
            lt.Reload()
            t.Commit()
            ok.append(name)
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            fail.append((name, str(ex)))

        if delay_seconds and delay_seconds > 0:
            time.sleep(delay_seconds)

    return (ok, fail)


# ------------------------------------------------------------------
# Sync / close
# ------------------------------------------------------------------

def sync_with_central(doc):
    """SynchronizeWithCentral + full relinquish."""
    tcc = TransactWithCentralOptions()
    sync = SynchronizeWithCentralOptions()

    try:
        sync.SaveLocalBefore = True
    except Exception:
        pass
    try:
        sync.SaveLocalAfter = True
    except Exception:
        pass

    rel = RelinquishOptions(True)
    rel.UserWorksets = True
    rel.ViewWorksets = True
    rel.FamilyWorksets = True
    rel.StandardWorksets = True
    rel.CheckedOutElements = True
    sync.SetRelinquishOptions(rel)

    doc.SynchronizeWithCentral(tcc, sync)


def safe_close(doc):
    """Close the document without saving local changes."""
    try:
        doc.Close(False)
    except Exception:
        pass


# ------------------------------------------------------------------
# Full open-sync-close pipeline for a single model
# ------------------------------------------------------------------

def open_sync_close(app, region, project_guid, model_guid,
                    open_worksets=True,
                    reload_links=True,
                    link_delay=2.0,
                    reload_latest=True,
                    sync_delay=1.0,
                    close_after=True):
    """Open a cloud model, optionally open worksets / reload links, sync, close.

    Returns a dict with keys: opened, worksets_opened, worksets_still_closed,
    links_ok, links_fail, synced, error.
    """
    result = {
        "opened": False,
        "worksets_opened": 0,
        "worksets_still_closed": 0,
        "links_ok": [],
        "links_fail": [],
        "synced": False,
        "error": None,
    }

    doc = None
    try:
        mp = build_cloud_model_path(region, project_guid, model_guid)
        open_opts = make_open_options(open_worksets)
        doc = app.OpenDocumentFile(mp, open_opts)
        result["opened"] = True

        # Worksets
        if open_worksets:
            wo, wsc = open_all_closed_user_worksets(doc)
            result["worksets_opened"] = wo
            result["worksets_still_closed"] = wsc

        # Links
        if reload_links:
            lok, lfail = reload_links_one_by_one(doc, delay_seconds=link_delay)
            result["links_ok"] = lok
            result["links_fail"] = lfail

        # ReloadLatest
        if reload_latest:
            try:
                doc.ReloadLatest()
            except Exception:
                pass

        # Settle
        if sync_delay and sync_delay > 0:
            time.sleep(sync_delay)

        # Sync
        sync_with_central(doc)
        result["synced"] = True

    except Exception as ex:
        result["error"] = str(ex)
    finally:
        if close_after and doc:
            safe_close(doc)

    return result
