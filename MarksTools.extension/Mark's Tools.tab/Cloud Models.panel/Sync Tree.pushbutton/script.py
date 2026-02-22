# -*- coding: utf-8 -*-
"""Dependency-aware cloud model sync wizard.

Selects a root model, discovers all linked children via BFS,
then syncs bottom-up (leaves first) through a WPF wizard.
"""
__title__ = "Sync Tree"
__doc__ = (
    "Discover the full link-dependency tree of an ACC cloud model "
    "and sync every model bottom-up (leaves first, root last)."
)
__author__ = "Mark"
__context__ = "active-doc"

import os
import threading

from pyrevit import revit, forms, script
import System
from System.Windows import Visibility
from System.Windows.Controls import (
    CheckBox as WpfCheckBox,
    TextBlock as WpfTextBlock,
)
from System.Windows.Media import SolidColorBrush
from System.Windows import Thickness

from guid_extractor import extract_cloud_ids, safe_str
from dependency_tree import build_dependency_tree, sync_order, format_tree
from cloud_helpers import open_sync_close

logger = script.get_logger()

# ──────────────────────────────────────────────────────────────
# Colour helpers (match XAML resources)
# ──────────────────────────────────────────────────────────────
_CLR_SUCCESS = SolidColorBrush(System.Windows.Media.Color.FromRgb(78, 201, 176))
_CLR_ERROR   = SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 76, 76))
_CLR_FG      = SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220))
_CLR_DIM     = SolidColorBrush(System.Windows.Media.Color.FromRgb(160, 160, 160))


# ══════════════════════════════════════════════════════════════
# Wizard Window
# ══════════════════════════════════════════════════════════════
class SyncWizard(forms.WPFWindow):

    def __init__(self):
        xaml = os.path.join(os.path.dirname(__file__), "SyncWizard.xaml")
        forms.WPFWindow.__init__(self, xaml)

        self._step = 0
        self._adjacency = {}
        self._model_info = {}
        self._ordered = []
        self._tree_checks = {}           # model_guid -> WPF CheckBox
        self._root_region = ""
        self._root_project = ""
        self._root_model = ""
        self._root_name = ""

        # Wire manual-panel toggle
        self.rb_manual.Checked  += self._toggle_manual
        self.rb_active.Checked  += self._toggle_manual

        self._populate_active_info()
        self._show_step(0)

    # ── Active-document info ──────────────────────────────────
    def _populate_active_info(self):
        doc = revit.doc
        if not doc:
            return
        try:
            region, pg, mg = extract_cloud_ids(doc)
            self.lbl_active_title.Text   = doc.Title
            self.lbl_active_region.Text  = "Region: " + region
            self.lbl_active_project.Text = "Project GUID: " + pg
            self.lbl_active_model.Text   = "Model GUID: " + mg
            self._root_region  = region
            self._root_project = pg
            self._root_model   = mg
            self._root_name    = doc.Title
        except Exception as ex:
            self.lbl_active_title.Text = "Could not read IDs: " + safe_str(ex)

    def _toggle_manual(self, sender, args):
        if self.rb_manual.IsChecked:
            self.manual_panel.Visibility = Visibility.Visible
            self.active_info_panel.Visibility = Visibility.Collapsed
        else:
            self.manual_panel.Visibility = Visibility.Collapsed
            self.active_info_panel.Visibility = Visibility.Visible

    # ── Step navigation ───────────────────────────────────────
    def _show_step(self, step):
        self._step = step
        for p in (self.step1_panel, self.step2_panel,
                  self.step3_panel, self.step4_panel):
            p.Visibility = Visibility.Collapsed

        panels = [self.step1_panel, self.step2_panel,
                  self.step3_panel, self.step4_panel]
        panels[step].Visibility = Visibility.Visible

        self.btn_back.Visibility  = Visibility.Visible if step > 0 and step < 3 else Visibility.Collapsed
        self.btn_next.Visibility  = Visibility.Visible if step < 2 else Visibility.Collapsed
        self.btn_run.Visibility   = Visibility.Visible if step == 2 else Visibility.Collapsed
        self.btn_close.Visibility = Visibility.Visible if step == 3 else Visibility.Collapsed

    # ── Button handlers ───────────────────────────────────────
    def btn_back_click(self, sender, args):
        if self._step > 0:
            self._show_step(self._step - 1)

    def btn_next_click(self, sender, args):
        if self._step == 0:
            if not self._read_root_ids():
                return
            self._show_step(1)
            self._start_discovery()
        elif self._step == 1:
            self._show_step(2)
            cnt = sum(1 for g in self._ordered if self._tree_checks.get(g, None) and self._tree_checks[g].IsChecked)
            self.lbl_model_count.Text = "{0} model(s) selected for sync.".format(cnt)

    def btn_run_click(self, sender, args):
        self._show_step(3)
        self._start_sync()

    def btn_close_click(self, sender, args):
        self.Close()

    # ── Step 1: read root identifiers ─────────────────────────
    def _read_root_ids(self):
        if self.rb_manual.IsChecked:
            r  = self.tb_region.Text.strip()
            pg = self.tb_project_guid.Text.strip()
            mg = self.tb_model_guid.Text.strip()
            if not r or not pg or not mg:
                forms.alert("Please fill in Region, Project GUID, and Model GUID.")
                return False
            self._root_region  = r
            self._root_project = pg
            self._root_model   = mg
            self._root_name    = self.tb_model_name.Text.strip() or "ROOT"
        else:
            if not self._root_model:
                forms.alert("Active document does not appear to be a cloud model.")
                return False
        return True

    # ── Step 2: BFS discovery (background thread) ─────────────
    def _start_discovery(self):
        self.pb_scan.IsIndeterminate = True
        self.lbl_scan_status.Text = "Opening linked models (detached) to discover tree…"
        self.btn_next.IsEnabled = False

        t = threading.Thread(target=self._discover_worker)
        t.IsBackground = True
        t.Start()

    def _discover_worker(self):
        try:
            app = revit.HOST_APP.app

            def progress_cb(msg):
                self.Dispatcher.Invoke(
                    lambda: setattr(self.lbl_scan_status, "Text", msg)
                )

            adj, info = build_dependency_tree(
                app,
                self._root_region,
                self._root_project,
                self._root_model,
                root_name=self._root_name,
                progress_callback=progress_cb,
            )
            ordered = sync_order(adj, info)

            self.Dispatcher.Invoke(
                lambda: self._on_discovery_done(adj, info, ordered)
            )
        except Exception as ex:
            self.Dispatcher.Invoke(
                lambda: self._on_discovery_error(str(ex))
            )

    def _on_discovery_done(self, adj, info, ordered):
        self._adjacency  = adj
        self._model_info = info
        self._ordered    = ordered

        self.pb_scan.IsIndeterminate = False
        self.pb_scan.Value = 100
        self.lbl_scan_status.Text = "Found {0} model(s).".format(len(ordered))
        self.btn_next.IsEnabled = True

        # Build tree UI with checkboxes
        self.tree_panel.Children.Clear()
        self._tree_checks.clear()

        tree_lines = format_tree(adj, info, self._root_model)

        for indent, guid, name in tree_lines:
            sp = System.Windows.Controls.StackPanel()
            sp.Orientation = System.Windows.Controls.Orientation.Horizontal
            sp.Margin = Thickness(indent * 20, 2, 0, 2)

            cb = WpfCheckBox()
            cb.IsChecked = True
            cb.Tag = guid
            sp.Children.Add(cb)

            lbl = WpfTextBlock()
            lbl.Text = name
            lbl.Foreground = _CLR_FG
            lbl.Margin = Thickness(4, 0, 8, 0)
            sp.Children.Add(lbl)

            guid_lbl = WpfTextBlock()
            guid_lbl.Text = guid[:13] + "…"
            guid_lbl.Foreground = _CLR_DIM
            guid_lbl.FontSize = 11
            sp.Children.Add(guid_lbl)

            self.tree_panel.Children.Add(sp)
            self._tree_checks[guid] = cb

    def _on_discovery_error(self, msg):
        self.pb_scan.IsIndeterminate = False
        self.lbl_scan_status.Text = "Discovery failed."
        self.lbl_scan_status.Foreground = _CLR_ERROR
        forms.alert("Tree discovery failed:\n\n" + msg)

    # ── Step 4: sync execution (background thread) ────────────
    def _start_sync(self):
        selected = [
            g for g in self._ordered
            if self._tree_checks.get(g) and self._tree_checks[g].IsChecked
        ]
        if not selected:
            forms.alert("No models selected.")
            return

        self.pb_sync.Maximum = len(selected)
        self.pb_sync.Value = 0
        self.tb_log.Text = ""
        self.lbl_sync_progress.Text = "Starting…"

        t = threading.Thread(target=self._sync_worker, args=(selected,))
        t.IsBackground = True
        t.Start()

    def _sync_worker(self, selected):
        app = revit.HOST_APP.app

        # Read options from UI (on UI thread)
        opts = {}
        def _read_opts():
            opts["open_ws"]     = self.cb_open_worksets.IsChecked
            opts["reload"]      = self.cb_reload_links.IsChecked
            opts["link_delay"]  = _safe_float(self.tb_link_delay.Text, 2.0)
            opts["latest"]      = self.cb_reload_latest.IsChecked
            opts["sync_delay"]  = _safe_float(self.tb_sync_delay.Text, 1.0)
            opts["close"]       = self.cb_close_after.IsChecked
        self.Dispatcher.Invoke(_read_opts)

        ok_count = 0
        fail_count = 0

        for idx, guid in enumerate(selected):
            info = self._model_info.get(guid, {})
            name = info.get("name", guid[:8])
            region = info.get("region", self._root_region)
            proj   = info.get("project_guid", self._root_project)

            self.Dispatcher.Invoke(
                lambda n=name, i=idx: self._ui_sync_progress(n, i, len(selected))
            )

            try:
                result = open_sync_close(
                    app, region, proj, guid,
                    open_worksets=opts["open_ws"],
                    reload_links=opts["reload"],
                    link_delay=opts["link_delay"],
                    reload_latest=opts["latest"],
                    sync_delay=opts["sync_delay"],
                    close_after=opts["close"],
                )

                if result["error"]:
                    fail_count += 1
                    self.Dispatcher.Invoke(
                        lambda n=name, e=result["error"]: self._log(
                            "FAIL  {0}: {1}".format(n, e), error=True
                        )
                    )
                else:
                    ok_count += 1
                    msg = "OK    {0}  (ws opened: {1}, links: {2}, link fails: {3})".format(
                        name,
                        result["worksets_opened"],
                        len(result["links_ok"]),
                        len(result["links_fail"]),
                    )
                    self.Dispatcher.Invoke(lambda m=msg: self._log(m))

            except Exception as ex:
                fail_count += 1
                self.Dispatcher.Invoke(
                    lambda n=name, e=str(ex): self._log(
                        "FAIL  {0}: {1}".format(n, e), error=True
                    )
                )

        summary = "Done. {0} succeeded, {1} failed.".format(ok_count, fail_count)
        self.Dispatcher.Invoke(lambda: self._on_sync_done(summary))

    # ── UI helpers (must run on dispatcher thread) ─────────────
    def _ui_sync_progress(self, name, idx, total):
        self.lbl_sync_progress.Text = "Syncing {0}/{1}: {2}".format(idx + 1, total, name)
        self.pb_sync.Value = idx

    def _log(self, msg, error=False):
        self.tb_log.Text += msg + "\n"
        if error:
            logger.warning(msg)

    def _on_sync_done(self, summary):
        self.lbl_sync_progress.Text = summary
        self.pb_sync.Value = self.pb_sync.Maximum
        self.btn_close.Visibility = Visibility.Visible


# ──────────────────────────────────────────────────────────────
# Utility
# ──────────────────────────────────────────────────────────────
def _safe_float(text, default):
    try:
        return float(text)
    except Exception:
        return default


# ══════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════
wizard = SyncWizard()
wizard.ShowDialog()
