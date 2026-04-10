# -*- coding: utf-8 -*-
"""
Property Editor Dialog for external classifications.
"""

from pyrevit import forms
import copy
import properties_editor as pe


class PropertyEditorDialog(forms.WPFWindow):
    """Dialog to configure properties for classified elements."""

    def __init__(self, doc, element_ids, class_data, dict_data):
        import os

        self._ready = False
        xaml_path = os.path.join(os.path.dirname(__file__), "properties_dialog.xaml")
        forms.WPFWindow.__init__(self, xaml_path)

        self.doc = doc
        self.element_ids = list(element_ids or [])
        self.class_data = class_data or {}
        self.dict_data = dict_data or {}
        self.class_uri = self.class_data.get("uri", "") or self.class_data.get("classUri", "")
        self.result = None

        self._base_property_sets = pe.build_property_tree(self.class_data)
        self._element_infos = self._build_element_infos()
        self._element_states = {}
        for info in self._element_infos:
            self._element_states[info["id_int"]] = self._load_initial_state_for_element(info.get("eid"))

        if self._element_infos:
            first_id = self._element_infos[0]["id_int"]
            self._shared_state = copy.deepcopy(self._element_states.get(first_id, self._base_property_sets))
        else:
            self._shared_state = copy.deepcopy(self._base_property_sets)

        self._active_element_id_int = self._element_infos[0]["id_int"] if self._element_infos else None

        self._prop_controls = {}   # key -> {"cb":..., "input":..., "prop":..., "pset":...}
        self._pset_groups = {}     # pset_name -> {"parent":..., "children":[], "counter":...}
        self._suspend_events = False

        self._populate_element_selector()
        self._update_mode_visibility()
        self._render_current_state()
        self._ready = True

    def _build_element_infos(self):
        infos = []
        for eid in self.element_ids:
            try:
                elem = self.doc.GetElement(eid)
                if elem and elem.IsValidObject:
                    try:
                        elem_name = elem.Name or "Element"
                    except Exception:
                        elem_name = "Element"
                else:
                    elem_name = "Element"
                infos.append({
                    "eid": eid,
                    "id_int": eid.IntegerValue,
                    "label": u"{} [ID: {}]".format(elem_name, eid.IntegerValue)
                })
            except Exception:
                pass
        return infos

    def _load_initial_state_for_element(self, eid):
        """Build editable state by merging stored values into the class property template."""
        state = copy.deepcopy(self._base_property_sets)
        try:
            elem = self.doc.GetElement(eid)
            if not elem or not elem.IsValidObject or not self.class_uri:
                return state

            stored = pe.load_class_properties(elem, self.class_uri)
            stored_sets = list((stored or {}).get("property_sets", []) or [])
            if not stored_sets:
                return state

            stored_map = {}
            for pset in stored_sets:
                sname = pset.get("name", "")
                for prop in list(pset.get("properties", []) or []):
                    key_uri = prop.get("uri", "") or ""
                    key_name = prop.get("name", "") or ""
                    key = (sname, key_uri, key_name)
                    stored_map[key] = prop

            for pset in state:
                pname = pset.get("name", "")
                for prop in list(pset.get("properties", []) or []):
                    key = (pname, prop.get("uri", "") or "", prop.get("name", "") or "")
                    sprop = stored_map.get(key)
                    if sprop is None:
                        continue
                    prop["enabled"] = bool(sprop.get("enabled", False))
                    prop["value"] = sprop.get("value", "") or ""

        except Exception:
            pass
        return state

    def _is_per_element_mode(self):
        return bool(self.ApplyPerElementRadio.IsChecked)

    def _populate_element_selector(self):
        import clr
        clr.AddReference('PresentationFramework')
        clr.AddReference('PresentationCore')
        clr.AddReference('WindowsBase')
        from System.Windows.Controls import ComboBoxItem

        self._suspend_events = True
        self.ElementSelectorCombo.Items.Clear()
        for info in self._element_infos:
            item = ComboBoxItem()
            item.Content = info["label"]
            item.Tag = info["id_int"]
            self.ElementSelectorCombo.Items.Add(item)

        if self.ElementSelectorCombo.Items.Count > 0:
            self.ElementSelectorCombo.SelectedIndex = 0
        self._suspend_events = False

    def _update_mode_visibility(self):
        import clr
        clr.AddReference('PresentationFramework')
        clr.AddReference('PresentationCore')
        clr.AddReference('WindowsBase')
        from System.Windows import Visibility

        per_mode = self._is_per_element_mode()
        self.ElementSelectorPanel.Visibility = Visibility.Visible if per_mode else Visibility.Collapsed
        nav_enabled = per_mode and self.ElementSelectorCombo.Items.Count > 1
        self.PrevElementBtn.IsEnabled = nav_enabled
        self.NextElementBtn.IsEnabled = nav_enabled

    def _get_selected_element_id_int(self):
        item = self.ElementSelectorCombo.SelectedItem
        if item and hasattr(item, "Tag"):
            return item.Tag
        return self._active_element_id_int

    def _get_state_for_editing(self):
        if self._is_per_element_mode():
            key = self._active_element_id_int
            if key in self._element_states:
                return self._element_states[key]
            return copy.deepcopy(self._base_property_sets)
        return self._shared_state

    def _render_current_state(self):
        state = self._get_state_for_editing()
        self._populate_properties_panel(state)
        self._update_status()

    def _populate_properties_panel(self, property_sets):
        import clr
        clr.AddReference('PresentationFramework')
        clr.AddReference('PresentationCore')
        clr.AddReference('WindowsBase')
        from System.Windows.Controls import (
            CheckBox, TextBlock, TextBox, ComboBox, StackPanel, Border,
            Expander, Grid, ColumnDefinition
        )
        from System.Windows.Controls import Orientation
        from System.Windows import VerticalAlignment, Thickness, CornerRadius, FontWeights, TextWrapping
        from System.Windows import GridLength, GridUnitType
        from System.Windows.Media import Brushes

        self.PropertySetsPanel.Children.Clear()
        self._prop_controls = {}
        self._pset_groups = {}

        for pset in property_sets:
            pset_name = pset.get("name", "Other")
            props = list(pset.get("properties", []) or [])

            container = Border()
            container.BorderBrush = Brushes.Gainsboro
            container.BorderThickness = Thickness(1)
            container.CornerRadius = CornerRadius(4)
            container.Margin = Thickness(0, 0, 0, 8)
            container.Padding = Thickness(6)

            expander = Expander()
            expander.IsExpanded = True

            header_grid = Grid()
            c0 = ColumnDefinition()
            c0.Width = GridLength.Auto
            c1 = ColumnDefinition()
            c1.Width = GridLength(1, GridUnitType.Star)
            c2 = ColumnDefinition()
            c2.Width = GridLength.Auto
            header_grid.ColumnDefinitions.Add(c0)
            header_grid.ColumnDefinitions.Add(c1)
            header_grid.ColumnDefinitions.Add(c2)

            pset_cb = CheckBox()
            pset_cb.Tag = pset_name
            pset_cb.VerticalAlignment = VerticalAlignment.Center
            pset_cb.IsThreeState = True
            pset_cb.Checked += self._pset_checkbox_changed
            pset_cb.Unchecked += self._pset_checkbox_changed
            pset_cb.Indeterminate += self._pset_checkbox_changed
            Grid.SetColumn(pset_cb, 0)
            header_grid.Children.Add(pset_cb)

            pset_title = TextBlock()
            pset_title.Text = u"{}".format(pset_name)
            pset_title.FontWeight = FontWeights.SemiBold
            pset_title.Margin = Thickness(8, 0, 0, 0)
            pset_title.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(pset_title, 1)
            header_grid.Children.Add(pset_title)

            pset_counter = TextBlock()
            pset_counter.Foreground = Brushes.DimGray
            pset_counter.VerticalAlignment = VerticalAlignment.Center
            pset_counter.Margin = Thickness(14, 0, 0, 0)
            Grid.SetColumn(pset_counter, 2)
            header_grid.Children.Add(pset_counter)

            expander.Header = header_grid

            body = StackPanel()
            body.Orientation = Orientation.Vertical
            body.Margin = Thickness(0, 8, 0, 0)

            self._pset_groups[pset_name] = {"parent": pset_cb, "children": [], "counter": pset_counter}

            for idx, prop in enumerate(props):
                row = Grid()
                row.Margin = Thickness(0, 0, 0, 6)

                rc0 = ColumnDefinition()
                rc0.Width = GridLength.Auto
                rc1 = ColumnDefinition()
                rc1.Width = GridLength(1, GridUnitType.Star)
                rc2 = ColumnDefinition()
                rc2.Width = GridLength(300)
                rc3 = ColumnDefinition()
                rc3.Width = GridLength(100)
                row.ColumnDefinitions.Add(rc0)
                row.ColumnDefinitions.Add(rc1)
                row.ColumnDefinitions.Add(rc2)
                row.ColumnDefinitions.Add(rc3)

                prop_cb = CheckBox()
                prop_cb.IsChecked = bool(prop.get("enabled", False))
                prop_cb.VerticalAlignment = VerticalAlignment.Top
                prop_cb.Margin = Thickness(0, 3, 0, 0)
                prop_cb.Checked += self._prop_checkbox_changed
                prop_cb.Unchecked += self._prop_checkbox_changed
                Grid.SetColumn(prop_cb, 0)
                row.Children.Add(prop_cb)

                prop_name = TextBlock()
                prop_name.Text = prop.get("name", "")
                prop_name.TextWrapping = TextWrapping.Wrap
                prop_name.Margin = Thickness(8, 0, 10, 0)
                prop_name.VerticalAlignment = VerticalAlignment.Center
                Grid.SetColumn(prop_name, 1)
                row.Children.Add(prop_name)

                raw_dtype = prop.get("dataType", "")
                try:
                    dtype = unicode(raw_dtype or "").strip().lower()
                except Exception:
                    dtype = u"{}".format(raw_dtype or "").strip().lower()

                if prop.get("allowedValues"):
                    input_ctrl = ComboBox()
                    input_ctrl.MinWidth = 220
                    input_ctrl.Margin = Thickness(0, 0, 10, 0)
                    values = []
                    for av in prop.get("allowedValues", []):
                        values.append(av.get("value") or av.get("code") or "")
                    for value in values:
                        input_ctrl.Items.Add(value)
                    current_value = prop.get("value", "")
                    if current_value:
                        input_ctrl.Text = current_value
                    elif input_ctrl.Items.Count > 0:
                        input_ctrl.SelectedIndex = 0
                elif dtype in ("boolean", "bool"):
                    input_ctrl = CheckBox()
                    input_ctrl.Margin = Thickness(0, 2, 10, 0)
                    input_ctrl.VerticalAlignment = VerticalAlignment.Center
                    input_ctrl.Content = "True"
                    raw_val = prop.get("value", "")
                    try:
                        raw_val_text = unicode(raw_val or "").strip().lower()
                    except Exception:
                        raw_val_text = u"{}".format(raw_val or "").strip().lower()
                    input_ctrl.IsChecked = raw_val_text in ("true", "1", "yes", "y")
                elif dtype in ("time", "date", "datetime"):
                    # Keep time-like fields as plain text for maximum flexibility
                    # (e.g. year only, custom format).
                    input_ctrl = TextBox()
                    input_ctrl.MinWidth = 220
                    input_ctrl.Margin = Thickness(0, 0, 10, 0)
                    input_ctrl.Text = prop.get("value", "") or prop.get("example", "") or ""
                else:
                    input_ctrl = TextBox()
                    input_ctrl.MinWidth = 220
                    input_ctrl.Margin = Thickness(0, 0, 10, 0)
                    input_ctrl.Text = prop.get("value", "") or prop.get("example", "") or ""


                # Clean up "none" or similar placeholder values
                try:
                    if isinstance(input_ctrl, TextBox):
                        current_text = (input_ctrl.Text or "").strip().lower()
                        if current_text in ("none", "null", "undefined", "na", "n/a"):
                            input_ctrl.Text = ""
                        else:
                            # Ensure value is safe unicode
                            input_ctrl.Text = unicode(input_ctrl.Text or "")
                except Exception:
                    if isinstance(input_ctrl, TextBox):
                        input_ctrl.Text = ""
                
                input_ctrl.IsEnabled = bool(prop_cb.IsChecked)
                Grid.SetColumn(input_ctrl, 2)
                row.Children.Add(input_ctrl)

                # Handle data type display - robust even if source is not a string.
                raw_type = prop.get("dataType", "")
                try:
                    data_type = unicode(raw_type or "").strip()
                except Exception:
                    try:
                        data_type = u"{}".format(raw_type).strip()
                    except Exception:
                        data_type = ""
                if data_type.lower() in ("", "none", "null"):
                    data_type = "undefined"
                # For undefined types, mention they're treated as string
                dtype_display = data_type if data_type != "undefined" else "undefined (string)"
                
                dtype_tb = TextBlock()
                dtype_tb.Text = u"{}".format(dtype_display)
                dtype_tb.Foreground = Brushes.Gray
                dtype_tb.FontSize = 10
                dtype_tb.VerticalAlignment = VerticalAlignment.Center
                Grid.SetColumn(dtype_tb, 3)
                row.Children.Add(dtype_tb)
                prop_key = u"{}::{}::{}".format(pset_name, prop.get("uri", ""), idx)
                self._prop_controls[prop_key] = {
                    "cb": prop_cb,
                    "input": input_ctrl,
                    "prop": prop,
                    "pset": pset_name
                }
                self._pset_groups[pset_name]["children"].append(prop_cb)

                body.Children.Add(row)

            expander.Content = body
            container.Child = expander
            self.PropertySetsPanel.Children.Add(container)
            self._refresh_pset_group_state(pset_name)

    def _capture_ui_to_state(self):
        for cfg in self._prop_controls.values():
            cb = cfg["cb"]
            input_ctrl = cfg["input"]
            prop = cfg["prop"]

            enabled = bool(cb.IsChecked)
            prop["enabled"] = enabled

            value = ""
            try:
                if hasattr(input_ctrl, "Content") and hasattr(input_ctrl, "IsChecked") and not hasattr(input_ctrl, "Text"):
                    value = "true" if bool(input_ctrl.IsChecked) else "false"
                elif hasattr(input_ctrl, "SelectedItem") and input_ctrl.SelectedItem is not None:
                    try:
                        value = unicode(input_ctrl.SelectedItem)
                    except Exception:
                        value = u"{}".format(input_ctrl.SelectedItem)
                elif hasattr(input_ctrl, "Text"):
                    value = input_ctrl.Text or ""
            except Exception:
                value = ""

            # Lightweight numeric normalization for typed fields
            try:
                raw_dtype = prop.get("dataType", "")
                dtype = unicode(raw_dtype or "").strip().lower()
            except Exception:
                dtype = u"{}".format(prop.get("dataType", "") or "").strip().lower()

            if dtype in ("integer", "int"):
                try:
                    iv = int((value or "").strip())
                    value = unicode(iv)
                except Exception:
                    value = ""
            elif dtype in ("real", "number", "float", "double", "decimal"):
                try:
                    txt = (value or "").strip().replace(",", ".")
                    fv = float(txt)
                    value = unicode(fv)
                except Exception:
                    value = ""

            prop["value"] = value

    def _refresh_pset_group_state(self, pset_name):
        grp = self._pset_groups.get(pset_name)
        if not grp:
            return

        children = grp.get("children", [])
        total = len(children)
        enabled = sum(1 for c in children if bool(c.IsChecked))

        parent = grp.get("parent")
        if total == 0:
            parent.IsChecked = False
        elif enabled == 0:
            parent.IsChecked = False
        elif enabled == total:
            parent.IsChecked = True
        else:
            parent.IsChecked = None

        counter = grp.get("counter")
        if counter:
            counter.Text = u"({} / {} selected)".format(enabled, total)

    def _set_control_enabled_for_checkbox(self, cb):
        for cfg in self._prop_controls.values():
            if cfg["cb"] is cb:
                cfg["input"].IsEnabled = bool(cb.IsChecked)
                break

    def _pset_checkbox_changed(self, sender, args):
        if not getattr(self, "_ready", False):
            return
        if self._suspend_events:
            return
        pset_name = sender.Tag if hasattr(sender, "Tag") else None
        grp = self._pset_groups.get(pset_name)
        if not grp:
            return

        target = sender.IsChecked
        if target is None:
            return

        self._suspend_events = True
        try:
            for child in grp.get("children", []):
                child.IsChecked = bool(target)
                self._set_control_enabled_for_checkbox(child)
        finally:
            self._suspend_events = False

        self._refresh_pset_group_state(pset_name)
        self._update_status()

    def _prop_checkbox_changed(self, sender, args):
        if not getattr(self, "_ready", False):
            return
        if self._suspend_events:
            return
        self._set_control_enabled_for_checkbox(sender)
        for pset_name, grp in self._pset_groups.items():
            if sender in grp.get("children", []):
                self._refresh_pset_group_state(pset_name)
                break
        self._update_status()

    def _update_status(self):
        enabled_count = sum(1 for cfg in self._prop_controls.values() if bool(cfg["cb"].IsChecked))
        total_count = len(self._prop_controls)

        if self._is_per_element_mode():
            item = self.ElementSelectorCombo.SelectedItem
            label = item.Content if item and hasattr(item, "Content") else "(no element selected)"
            self.StatusText.Text = u"Individual mode | {} | Selected: {} / {} properties".format(
                label, enabled_count, total_count
            )
        else:
            self.StatusText.Text = u"All-elements mode | Selected: {} / {} properties".format(enabled_count, total_count)

    def ApplyMode_Changed(self, sender, args):
        if not getattr(self, "_ready", False):
            return
        if self._suspend_events:
            return
        self._capture_ui_to_state()
        self._update_mode_visibility()
        if self._is_per_element_mode() and self.ElementSelectorCombo.SelectedIndex < 0 and self.ElementSelectorCombo.Items.Count > 0:
            self.ElementSelectorCombo.SelectedIndex = 0
        self._active_element_id_int = self._get_selected_element_id_int()
        self._render_current_state()

    def ElementSelector_Changed(self, sender, args):
        if not getattr(self, "_ready", False):
            return
        if self._suspend_events:
            return
        self._capture_ui_to_state()
        self._active_element_id_int = self._get_selected_element_id_int()
        self._render_current_state()

    def PrevElement_Click(self, sender, args):
        if not getattr(self, "_ready", False):
            return
        if self.ElementSelectorCombo.Items.Count <= 1:
            return
        idx = self.ElementSelectorCombo.SelectedIndex
        if idx > 0:
            self.ElementSelectorCombo.SelectedIndex = idx - 1

    def NextElement_Click(self, sender, args):
        if not getattr(self, "_ready", False):
            return
        count = self.ElementSelectorCombo.Items.Count
        if count <= 1:
            return
        idx = self.ElementSelectorCombo.SelectedIndex
        if idx < count - 1:
            self.ElementSelectorCombo.SelectedIndex = idx + 1

    def _extract_enabled_property_sets(self, property_sets):
        enabled_sets = []
        for pset in property_sets:
            enabled_props = []
            for prop in pset.get("properties", []):
                if prop.get("enabled"):
                    enabled_props.append(prop)
            if enabled_props:
                enabled_sets.append({
                    "name": pset.get("name", "Other"),
                    "description": pset.get("description", ""),
                    "properties": enabled_props
                })
        return enabled_sets

    def Save_Click(self, sender, args):
        try:
            from Autodesk.Revit.DB import Transaction as TX

            self._capture_ui_to_state()

            if not self._element_infos:
                forms.alert("No valid selected elements found for property assignment.", title="Error")
                return

            if not self.class_uri:
                forms.alert("No class URI available. Properties can not be saved.", title="Error")
                return

            mode_all = bool(self.ApplyToAllRadio.IsChecked)

            if mode_all:
                enabled_sets = self._extract_enabled_property_sets(self._shared_state)
                payload_by_element = {}
                data = pe.serialize_properties_for_storage(enabled_sets) if enabled_sets else None
                for info in self._element_infos:
                    payload_by_element[info["id_int"]] = data
            else:
                payload_by_element = {}
                for info in self._element_infos:
                    state = self._element_states.get(info["id_int"], [])
                    enabled_sets = self._extract_enabled_property_sets(state)
                    if enabled_sets:
                        payload_by_element[info["id_int"]] = pe.serialize_properties_for_storage(enabled_sets)
                    else:
                        payload_by_element[info["id_int"]] = None

            # Ensure schema exists before entering transaction.
            if not pe.create_property_schema():
                forms.alert("Property schema could not be created.", title="Error")
                return

            tx = TX(self.doc, "Save Classification Properties")
            tx.Start()
            persisted_count = 0
            target_count = 0
            try:
                for info in self._element_infos:
                    elem = self.doc.GetElement(info["eid"])
                    if not elem or not elem.IsValidObject:
                        continue

                    payload = payload_by_element.get(info["id_int"])
                    if payload:
                        target_count += 1
                        if pe.store_class_properties(elem, self.class_uri, payload):
                            persisted_count += 1
                    else:
                        pe.delete_class_properties(elem, self.class_uri)
                tx.Commit()
            except Exception:
                try:
                    tx.RollBack()
                except Exception:
                    pass
                raise

            if target_count > 0 and persisted_count == 0:
                detail = pe.get_last_store_error()
                forms.alert(
                    "No property data was persisted.\n"
                    "Class URI: {}\n"
                    "Target elements: {}\n"
                    "Last backend error: {}".format(
                        self.class_uri or "(empty)",
                        target_count,
                        detail or "(none)"
                    ),
                    title="Error"
                )
                return

            self.result = "saved"
            self.Close()

        except Exception as ex:
            forms.alert("Error saving properties: {}".format(str(ex)), title="Error")

    def Cancel_Click(self, sender, args):
        self.result = None
        self.Close()
