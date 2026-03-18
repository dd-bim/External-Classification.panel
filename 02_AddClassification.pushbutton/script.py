# -*- coding: utf-8 -*-
__title__ = "Add\nClassification"
__doc__ = "Add external classification to selected elements"

from pyrevit import forms


def main():
    import clr
    clr.AddReference('System')
    import System
    import os
    import json as _json
    from Autodesk.Revit.DB import Transaction, FilteredElementCollector, CategoryType, ElementId, BuiltInCategory
    from Autodesk.Revit.UI.Selection import ObjectType
    from Autodesk.Revit.Exceptions import OperationCanceledException
    from Autodesk.Revit.DB.ExtensibleStorage import (
        Schema, SchemaBuilder, Entity, AccessLevel
    )

    doc   = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument

    _URL_SCHEMA_GUID_STR = "12345678-1234-1234-1234-1234567890ab"
    _URL_FIELD           = "APIUrl"
    _CLS_SCHEMA_GUID_STR = "ABCDEF12-3456-7890-ABCD-EF1234567890"
    _CLS_SCHEMA_NAME     = "ExternalClassification"
    _CLS_MULTI_SCHEMA_GUID_STR = "ABCDEF12-3456-7890-ABCD-EF1234567891"
    _CLS_MULTI_SCHEMA_NAME     = "ExternalClassificationMulti"

    def normalize_text(value):
        return (value or "").strip().lower()

    def get_stored_url():
        schema = Schema.Lookup(System.Guid(_URL_SCHEMA_GUID_STR))
        if schema:
            entity = doc.ProjectInformation.GetEntity(schema)
            if entity.IsValid():
                field = schema.GetField(_URL_FIELD)
                stored = entity.Get[System.String](field)
                if stored:
                    return stored
        # Fallback to default API if nothing stored
        return "https://api.bsdd.buildingsmart.org"

    def get_or_create_cls_schema():
        guid = System.Guid(_CLS_SCHEMA_GUID_STR)
        schema = Schema.Lookup(guid)
        if schema:
            return schema
        sb = SchemaBuilder(guid)
        sb.SetSchemaName(_CLS_SCHEMA_NAME)
        sb.SetReadAccessLevel(AccessLevel.Public)
        sb.SetWriteAccessLevel(AccessLevel.Public)
        for fname in ("Code", "Name", "ClassUri", "DictionaryName", "DictionaryUri"):
            sb.AddSimpleField(fname, System.String)
        return sb.Finish()

    def get_or_create_cls_multi_schema():
        guid = System.Guid(_CLS_MULTI_SCHEMA_GUID_STR)
        schema = Schema.Lookup(guid)
        if schema:
            return schema
        sb = SchemaBuilder(guid)
        sb.SetSchemaName(_CLS_MULTI_SCHEMA_NAME)
        sb.SetReadAccessLevel(AccessLevel.Public)
        sb.SetWriteAccessLevel(AccessLevel.Public)
        sb.AddSimpleField("ItemsJson", System.String)
        return sb.Finish()

    def load_element_classifications(elem, legacy_schema=None, legacy_fields=None, multi_schema=None, multi_items_field=None):
        items = []

        try:
            if multi_schema and multi_items_field:
                ment = elem.GetEntity(multi_schema)
                if ment and ment.IsValid():
                    raw = ment.Get[System.String](multi_items_field) or ""
                    if raw.strip():
                        payload = _json.loads(raw)
                        if isinstance(payload, list):
                            for p in payload:
                                if not isinstance(p, dict):
                                    continue
                                items.append({
                                    "code": (p.get("code") or "").strip(),
                                    "name": (p.get("name") or "").strip(),
                                    "class_uri": (p.get("class_uri") or "").strip(),
                                    "dict_name": (p.get("dict_name") or "").strip(),
                                    "dict_uri": (p.get("dict_uri") or "").strip(),
                                })
        except Exception:
            items = []

        if items:
            return items

        try:
            if legacy_schema and legacy_fields:
                ent = elem.GetEntity(legacy_schema)
                if ent and ent.IsValid():
                    code = ent.Get[System.String](legacy_fields["Code"]) or ""
                    if code:
                        items.append({
                            "code": code,
                            "name": ent.Get[System.String](legacy_fields["Name"]) or "",
                            "class_uri": ent.Get[System.String](legacy_fields["ClassUri"]) or "",
                            "dict_name": ent.Get[System.String](legacy_fields["DictionaryName"]) or "",
                            "dict_uri": ent.Get[System.String](legacy_fields["DictionaryUri"]) or "",
                        })
        except Exception:
            pass

        return items

    def store_element_classifications(elem, items, legacy_schema=None, legacy_fields=None, multi_schema=None, multi_items_field=None):
        clean_items = []
        seen = set()
        for it in items:
            d_name = (it.get("dict_name") or "").strip()
            d_uri = (it.get("dict_uri") or "").strip()
            code = (it.get("code") or "").strip()
            name = (it.get("name") or "").strip()
            class_uri = (it.get("class_uri") or "").strip()
            if not code or not d_name:
                continue
            key = (normalize_text(d_name), normalize_text(d_uri), normalize_text(code), normalize_text(name))
            if key in seen:
                continue
            seen.add(key)
            clean_items.append({
                "code": code,
                "name": name,
                "class_uri": class_uri,
                "dict_name": d_name,
                "dict_uri": d_uri,
            })

        if multi_schema and multi_items_field:
            ment = Entity(multi_schema)
            ment.Set[System.String](multi_items_field, _json.dumps(clean_items))
            elem.SetEntity(ment)

        # Keep legacy schema in sync with the newest entry for backward compatibility.
        if legacy_schema and legacy_fields and clean_items:
            latest = clean_items[-1]
            ent = Entity(legacy_schema)
            ent.Set[System.String](legacy_fields["Code"], latest.get("code") or "")
            ent.Set[System.String](legacy_fields["Name"], latest.get("name") or "")
            ent.Set[System.String](legacy_fields["ClassUri"], latest.get("class_uri") or "")
            ent.Set[System.String](legacy_fields["DictionaryName"], latest.get("dict_name") or "")
            ent.Set[System.String](legacy_fields["DictionaryUri"], latest.get("dict_uri") or "")
            elem.SetEntity(ent)

    def store_classification(elem, code, name, class_uri, dict_name, dict_uri, schema=None, fields=None):
        schema = schema or get_or_create_cls_schema()
        fields = fields or {
            "Code": schema.GetField("Code"),
            "Name": schema.GetField("Name"),
            "ClassUri": schema.GetField("ClassUri"),
            "DictionaryName": schema.GetField("DictionaryName"),
            "DictionaryUri": schema.GetField("DictionaryUri"),
        }
        entity = Entity(schema)
        entity.Set[System.String](fields["Code"], code or "")
        entity.Set[System.String](fields["Name"], name or "")
        entity.Set[System.String](fields["ClassUri"], class_uri or "")
        entity.Set[System.String](fields["DictionaryName"], dict_name or "")
        entity.Set[System.String](fields["DictionaryUri"], dict_uri or "")
        elem.SetEntity(entity)
        # Keep Add Classification stable by writing only ExtensibleStorage.
        # IFC/shared parameters are generated from stored classification during export.

    def is_camera_element(elem):
        """Return True for Revit camera elements (3D view cameras)."""
        try:
            cat = elem.Category
            if not cat:
                return False
            return cat.Id.IntegerValue == int(BuiltInCategory.OST_Cameras)
        except Exception:
            return False

    def apply_classification_to_elements(element_ids, cls_data, dict_data):
        dict_name = (dict_data or {}).get("name", "")
        dict_uri  = (dict_data or {}).get("uri", "")
        code      = (cls_data or {}).get("code", (cls_data or {}).get("referenceCode", ""))
        name      = (cls_data or {}).get("name", "")
        class_uri = (cls_data or {}).get("classUri", (cls_data or {}).get("uri", ""))

        elems = []
        for eid in element_ids:
            try:
                elem = doc.GetElement(eid)
                if elem and elem.IsValidObject:
                    if is_camera_element(elem):
                        continue
                    elems.append(elem)
            except Exception:
                pass

        success, errors = 0, 0
        legacy_schema = get_or_create_cls_schema()
        legacy_fields = {
            "Code": legacy_schema.GetField("Code"),
            "Name": legacy_schema.GetField("Name"),
            "ClassUri": legacy_schema.GetField("ClassUri"),
            "DictionaryName": legacy_schema.GetField("DictionaryName"),
            "DictionaryUri": legacy_schema.GetField("DictionaryUri"),
        }
        multi_schema = get_or_create_cls_multi_schema()
        multi_items_field = multi_schema.GetField("ItemsJson") if multi_schema else None

        target_dict_uri_norm = normalize_text(dict_uri)
        target_dict_name_norm = normalize_text(dict_name)

        def same_system(item):
            item_uri_norm = normalize_text(item.get("dict_uri", ""))
            item_name_norm = normalize_text(item.get("dict_name", ""))
            if target_dict_uri_norm and item_uri_norm:
                return item_uri_norm == target_dict_uri_norm
            return item_name_norm == target_dict_name_norm

        tx_write = Transaction(doc, "Set External Classification")
        try:
            tx_write.Start()
            for elem in elems:
                try:
                    existing = load_element_classifications(
                        elem,
                        legacy_schema=legacy_schema,
                        legacy_fields=legacy_fields,
                        multi_schema=multi_schema,
                        multi_items_field=multi_items_field
                    )
                    merged = [it for it in existing if not same_system(it)]
                    merged.append({
                        "code": code or "",
                        "name": name or "",
                        "class_uri": class_uri or "",
                        "dict_name": dict_name or "",
                        "dict_uri": dict_uri or "",
                    })
                    store_element_classifications(
                        elem,
                        merged,
                        legacy_schema=legacy_schema,
                        legacy_fields=legacy_fields,
                        multi_schema=multi_schema,
                        multi_items_field=multi_items_field
                    )
                    success += 1
                except Exception:
                    errors += 1
            tx_write.Commit()
            return success, errors, None
        except Exception as tx_ex:
            try:
                tx_write.RollBack()
            except Exception:
                pass
            return success, errors, tx_ex

    url = get_stored_url()
    # url now always has a value (default or stored)

    class ClassItem(object):
        def __init__(self, data):
            self.data            = data
            self.name            = data.get("name", "") or ""
            self.code            = data.get("code", data.get("referenceCode", "")) or ""
            self.descriptionPart = data.get("descriptionPart", "") or ""
        def ToString(self):
            return u"{} - {}".format(self.name, self.descriptionPart)
        def __str__(self):
            return self.ToString()

    class OverviewItem(object):
        def __init__(self, elem, classification_item=None, show_element_data=True):
            self.element_id = elem.Id
            has_classification = classification_item is not None
            self.is_classified_str = "True" if has_classification else "False"
            if has_classification and show_element_data:
                self.status = u"\u2713"
            elif has_classification:
                self.status = u"\u21B3"
            else:
                self.status = u"\u25cb"

            cat = elem.Category
            self.category = (cat.Name if cat else "") if show_element_data else ""
            try:
                base_name = elem.Name or str(elem.Id.IntegerValue)
            except Exception:
                base_name = str(elem.Id.IntegerValue)
            self.elem_name = base_name if show_element_data else ""

            self.classifications = ""
            self.classification_dict_name = ""
            self.classification_dict_uri = ""
            self.classification_code = ""
            self.classification_name = ""
            self.classification_uri = ""

            if has_classification:
                d_name = classification_item.get("dict_name", "") or "(dictionary)"
                d_uri = classification_item.get("dict_uri", "") or ""
                c_code = classification_item.get("code", "") or "(code)"
                c_name = classification_item.get("name", "") or ""
                c_uri = classification_item.get("class_uri", "") or ""

                if c_name:
                    self.classifications = u"{}: {} ({})".format(d_name, c_code, c_name)
                else:
                    self.classifications = u"{}: {}".format(d_name, c_code)

                self.classification_dict_name = d_name if d_name != "(dictionary)" else ""
                self.classification_dict_uri = d_uri
                self.classification_code = c_code if c_code != "(code)" else ""
                self.classification_name = c_name
                self.classification_uri = c_uri

    # Cache dictionary and class responses for the current command session
    # to avoid repeated API calls when the dialog is reopened.
    session_cache = {
        "dicts": None,
        "classes_by_uri": {}
    }

    class ClassificationDialog(forms.WPFWindow):
        def __init__(self, preselected_dict_uri=None):
            xaml_path = os.path.join(os.path.dirname(__file__), "dialog.xaml")
            forms.WPFWindow.__init__(self, xaml_path)
            self._url         = url.rstrip("/")
            self._all_classes = []
            self._dicts       = []
            self._loaded_dict_uri = ""
            self._pending_dict_uri = ""
            self._dropdown_open_selected_index = -1
            self._deferred_pick_request = None
            self._preselected_dict_uri = preselected_dict_uri or ""
            self._suspend_dict_selection_handler = False
            self.UrlDisplay.Text = self._url
            self._load_dictionaries()

        def _api_get(self, endpoint, params=None):
            """IronPython GET request to the external classification REST API."""
            import urllib2 as _ul, urllib as _ulbase, json as _json
            full = self._url + endpoint
            if params:
                safe = {}
                for k, v in params.items():
                    if isinstance(k, unicode): k = k.encode("utf-8")
                    if isinstance(v, unicode): v = v.encode("utf-8")
                    safe[k] = v
                full = full + "?" + _ulbase.urlencode(safe)
            resp = _ul.urlopen(full.encode("utf-8") if isinstance(full, unicode) else full, timeout=30)
            raw  = resp.read()
            return _json.loads(raw.decode("utf-8"))

        def _get(self, endpoint, params=None):
            return self._api_get(endpoint, params)

        def _set_status(self, msg):
            self.StatusText.Text = msg

        def _load_dictionaries(self):
            self._set_status("Loading dictionaries...")
            try:
                dicts = session_cache.get("dicts")
                if dicts is None:
                    result = self._get("/api/Dictionary/v1")
                    dicts = result.get("dictionaries", []) if isinstance(result, dict) else []
                    dicts = sorted(
                        dicts,
                        key=lambda d: ((d.get("name") if isinstance(d, dict) else str(d)) or u"").lower()
                    )
                    session_cache["dicts"] = dicts
                if not dicts:
                    self._set_status("No dictionaries found.")
                    return
                self._dicts = dicts
                self._suspend_dict_selection_handler = True
                self.DictionaryComboBox.Items.Clear()
                for d in self._dicts:
                    _dname = d.get("name") or str(d)
                    _dver = (d.get("version") or d.get("releaseDate") or "").strip()
                    _dlabel = u"{} ({})".format(_dname, _dver) if _dver else _dname
                    self.DictionaryComboBox.Items.Add(_dlabel)
                if self._preselected_dict_uri:
                    for i, d in enumerate(self._dicts):
                        if (d.get("uri", "") or "") == self._preselected_dict_uri:
                            self.DictionaryComboBox.SelectedIndex = i
                            self._suspend_dict_selection_handler = False
                            self._pending_dict_uri = d.get("uri", "") or ""
                            self.SearchBox.Text = ""
                            self._load_classes(self._pending_dict_uri)
                            break
                    else:
                        self._suspend_dict_selection_handler = False
                else:
                    self._suspend_dict_selection_handler = False
                self._set_status("Loaded {} dictionaries.".format(len(dicts)))
            except Exception as ex:
                self._set_status("Error: {}".format(ex))

        def _load_classes(self, dict_uri):
            self._set_status("Loading classes...")
            self.ClassListBox.Items.Clear()
            self._all_classes = []
            self._loaded_dict_uri = dict_uri or ""
            self.SearchBox.IsEnabled            = False
            self.AssignToSelectionBtn.IsEnabled = False
            self.SelectAndAssignBtn.IsEnabled   = False
            try:
                classes_by_uri = session_cache.setdefault("classes_by_uri", {})
                classes = classes_by_uri.get(dict_uri)
                if classes is None:
                    result = self._get("/api/Dictionary/v1/Classes", {"Uri": dict_uri})
                    classes = result.get("classes", []) if isinstance(result, dict) else []
                    classes.sort(key=lambda c: (c.get("name") or u"").lower())
                    classes_by_uri[dict_uri] = classes
                self._all_classes = classes
                self._populate_list(classes)
                self.SearchBox.IsEnabled = True
                self._set_status("Loaded {} classes.".format(len(classes)))
            except Exception as ex:
                self._set_status("Error: {}".format(ex))

        def _populate_list(self, classes):
            self.ClassListBox.Items.Clear()
            for c in classes:
                self.ClassListBox.Items.Add(ClassItem(c))

        def _get_selected_class(self):
            item = self.ClassListBox.SelectedItem
            return item.data if item else None

        def _apply_to_elements(self, element_ids, show_done_alert=True):
            cls_data = self._get_selected_class()
            if not cls_data:
                forms.alert("Please select a class first.", title="Notice")
                return
            idx       = self.DictionaryComboBox.SelectedIndex
            dict_data = self._dicts[idx] if (idx >= 0 and idx < len(self._dicts)) else {}

            success, errors, tx_ex = apply_classification_to_elements(element_ids, cls_data, dict_data)
            if tx_ex is not None:
                forms.alert(
                    "Classification write failed due to an unexpected Revit transaction error.",
                    title="Error"
                )
                return
            msg = "Classified {} element(s).".format(success)
            if errors:
                msg += " {} error(s).".format(errors)
            self._set_status(msg)
            if show_done_alert:
                forms.alert(msg, title="Done")

        def _remove_selected_classifications(self, selected_rows):
            legacy_schema = Schema.Lookup(System.Guid(_CLS_SCHEMA_GUID_STR))
            multi_schema = Schema.Lookup(System.Guid(_CLS_MULTI_SCHEMA_GUID_STR))
            legacy_fields = None
            if legacy_schema:
                legacy_fields = {
                    "Code": legacy_schema.GetField("Code"),
                    "Name": legacy_schema.GetField("Name"),
                    "ClassUri": legacy_schema.GetField("ClassUri"),
                    "DictionaryName": legacy_schema.GetField("DictionaryName"),
                    "DictionaryUri": legacy_schema.GetField("DictionaryUri"),
                }
            multi_items_field = multi_schema.GetField("ItemsJson") if multi_schema else None

            selected_by_elem = {}
            for row in selected_rows:
                try:
                    eid = row.element_id
                    if not eid:
                        continue
                    code = (row.classification_code or "").strip()
                    dname = (row.classification_dict_name or "").strip()
                    duri = (row.classification_dict_uri or "").strip()
                    cname = (row.classification_name or "").strip()
                    curi = (row.classification_uri or "").strip()
                    if not code and not dname:
                        continue
                    key = eid.IntegerValue
                    selected_by_elem.setdefault(key, []).append((
                        normalize_text(dname),
                        normalize_text(duri),
                        normalize_text(code),
                        normalize_text(cname),
                        normalize_text(curi),
                    ))
                except Exception:
                    pass

            if not selected_by_elem:
                forms.alert("Please select one or more classification rows in the overview.", title="Notice")
                return

            def _clear_all_named_params(elem, param_name):
                try:
                    for param in elem.Parameters:
                        try:
                            if param.Definition and param.Definition.Name == param_name and not param.IsReadOnly:
                                param.Set("")
                        except Exception:
                            pass
                except Exception:
                    pass

            def _matches_selector(item, selector):
                s_dname, s_duri, s_code, s_cname, s_curi = selector
                i_dname = normalize_text(item.get("dict_name", ""))
                i_duri = normalize_text(item.get("dict_uri", ""))
                i_code = normalize_text(item.get("code", ""))
                i_cname = normalize_text(item.get("name", ""))
                i_curi = normalize_text(item.get("class_uri", ""))

                if s_duri and i_duri and s_duri != i_duri:
                    return False
                if s_dname and s_dname != i_dname:
                    return False
                if s_code and s_code != i_code:
                    return False
                if s_cname and i_cname and s_cname != i_cname:
                    return False
                if s_curi and i_curi and s_curi != i_curi:
                    return False
                return True

            elements_changed, removed_lines, errors = 0, 0, 0
            with Transaction(doc, "Remove External Classification") as t:
                t.Start()
                for eid_int, selectors in selected_by_elem.items():
                    try:
                        elem = doc.GetElement(ElementId(eid_int))
                        if not elem or not elem.IsValidObject:
                            continue

                        existing = load_element_classifications(
                            elem,
                            legacy_schema=legacy_schema,
                            legacy_fields=legacy_fields,
                            multi_schema=multi_schema,
                            multi_items_field=multi_items_field
                        )
                        if not existing:
                            continue

                        keep_items = []
                        local_removed = 0
                        for it in existing:
                            matched = False
                            for sel in selectors:
                                if _matches_selector(it, sel):
                                    matched = True
                                    break
                            if matched:
                                local_removed += 1
                            else:
                                keep_items.append(it)

                        if local_removed == 0:
                            continue

                        if keep_items:
                            store_element_classifications(
                                elem,
                                keep_items,
                                legacy_schema=legacy_schema,
                                legacy_fields=legacy_fields,
                                multi_schema=multi_schema,
                                multi_items_field=multi_items_field
                            )
                        else:
                            if legacy_schema:
                                try:
                                    entity = elem.GetEntity(legacy_schema)
                                    if entity and entity.IsValid():
                                        elem.DeleteEntity(legacy_schema)
                                except Exception:
                                    pass

                            if multi_schema:
                                try:
                                    entity_multi = elem.GetEntity(multi_schema)
                                    if entity_multi and entity_multi.IsValid():
                                        elem.DeleteEntity(multi_schema)
                                except Exception:
                                    pass

                            for pname in (
                                "ClassificationCodePset",
                                "ClassificationName",
                                "ClassificationSystem",
                                "ClassificationCode",
                                "ClassificationCode(2)",
                                "ClassificationCode(3)",
                                "ClassificationCode(4)",
                                "ClassificationCode(5)",
                            ):
                                _clear_all_named_params(elem, pname)

                        elements_changed += 1
                        removed_lines += local_removed
                    except Exception:
                        errors += 1
                t.Commit()

            msg = "Removed {} classification line(s) on {} element(s).".format(removed_lines, elements_changed)
            if errors:
                msg += " {} error(s).".format(errors)
            self._set_status(msg)
            forms.alert(msg, title="Done")

        def DictionaryComboBox_SelectionChanged(self, sender, args):
            if self._suspend_dict_selection_handler:
                return
            idx = self.DictionaryComboBox.SelectedIndex
            if idx >= 0 and idx < len(self._dicts):
                self._pending_dict_uri = self._dicts[idx].get("uri", "") or ""

        def DictionaryComboBox_DropDownOpened(self, sender, args):
            self._dropdown_open_selected_index = self.DictionaryComboBox.SelectedIndex

        def DictionaryComboBox_DropDownClosed(self, sender, args):
            idx = self.DictionaryComboBox.SelectedIndex
            selection_changed = (idx != self._dropdown_open_selected_index)
            if selection_changed and self._pending_dict_uri and self._pending_dict_uri != self._loaded_dict_uri:
                self.SearchBox.Text = ""
                self._load_classes(self._pending_dict_uri)

        def SearchBox_TextChanged(self, sender, args):
            query = (self.SearchBox.Text or "").strip().lower()
            if not query:
                self._populate_list(self._all_classes)
                return
            filtered = [
                c for c in self._all_classes
                if query in (c.get("name", "") or "").lower()
                or query in (c.get("code", "") or "").lower()
                or query in (c.get("referenceCode", "") or "").lower()
            ]
            self._populate_list(filtered)

        def ClassListBox_SelectionChanged(self, sender, args):
            item = self.ClassListBox.SelectedItem
            can  = item is not None
            self.AssignToSelectionBtn.IsEnabled = can
            self.SelectAndAssignBtn.IsEnabled   = can
            if item:
                d = item.data
                self.DetailCode.Text = d.get("code", d.get("referenceCode", ""))
                self.DetailName.Text = d.get("name", "")
                self.DetailUri.Text  = d.get("classUri", d.get("uri", ""))
                self.ClassDetailPanel.Visibility = System.Windows.Visibility.Visible
            else:
                self.ClassDetailPanel.Visibility = System.Windows.Visibility.Collapsed

        def AssignToSelection_Click(self, sender, args):
            sel_ids = list(uidoc.Selection.GetElementIds())
            if not sel_ids:
                forms.alert("No elements selected.\nPlease select elements in Revit first.", title="Notice")
                return
            self._apply_to_elements(sel_ids)

        def SelectAndAssign_Click(self, sender, args):
            cls_data = self._get_selected_class()
            if not cls_data:
                forms.alert("Please select a class first.", title="Notice")
                return
            idx = self.DictionaryComboBox.SelectedIndex
            dict_data = self._dicts[idx] if (idx >= 0 and idx < len(self._dicts)) else {}
            self._deferred_pick_request = {
                "class_data": dict(cls_data),
                "dict_data": {
                    "name": dict_data.get("name", ""),
                    "uri": dict_data.get("uri", "")
                }
            }
            self.Close()

        def Close_Click(self, sender, args):
            self.Close()

        def Tab_SelectionChanged(self, sender, args):
            if args.OriginalSource != self.MainTabControl:
                return
            if self.MainTabControl.SelectedIndex == 1:
                self._load_overview()

        def RefreshOverview_Click(self, sender, args):
            self._load_overview()

        def _get_selected_overview_element_ids(self):
            element_ids = []
            seen = set()
            try:
                for item in self.OverviewList.SelectedItems:
                    try:
                        if item and hasattr(item, "element_id") and item.element_id:
                            key = item.element_id.IntegerValue
                            if key in seen:
                                continue
                            seen.add(key)
                            element_ids.append(item.element_id)
                    except Exception:
                        pass
            except Exception:
                pass
            return element_ids

        def SelectOverviewInModel_Click(self, sender, args):
            element_ids = self._get_selected_overview_element_ids()
            if not element_ids:
                forms.alert("Please select one or more rows in the overview first.", title="Notice")
                return
            try:
                id_list = System.Collections.Generic.List[ElementId]()
                for eid in element_ids:
                    id_list.Add(eid)
                uidoc.Selection.SetElementIds(id_list)
                uidoc.ShowElements(id_list)
                self._set_status("Selected {} element(s) in the model.".format(len(element_ids)))
            except Exception as ex:
                self._set_status("Selection error: {}".format(ex))

        def RemoveOverviewSelected_Click(self, sender, args):
            selected_rows = list(self.OverviewList.SelectedItems)
            if not selected_rows:
                forms.alert("Please select one or more rows in the overview first.", title="Notice")
                return
            self._remove_selected_classifications(selected_rows)
            self._load_overview()

        def _load_overview(self):
            self.OverviewList.Items.Clear()
            self.OverviewStats.Text = u""
            self._set_status(u"Loading overview...")
            legacy_schema = Schema.Lookup(System.Guid(_CLS_SCHEMA_GUID_STR))
            legacy_fields = None
            if legacy_schema:
                legacy_fields = {
                    "Code": legacy_schema.GetField("Code"),
                    "Name": legacy_schema.GetField("Name"),
                    "ClassUri": legacy_schema.GetField("ClassUri"),
                    "DictionaryName": legacy_schema.GetField("DictionaryName"),
                    "DictionaryUri": legacy_schema.GetField("DictionaryUri"),
                }
            multi_schema = Schema.Lookup(System.Guid(_CLS_MULTI_SCHEMA_GUID_STR))
            multi_items_field = multi_schema.GetField("ItemsJson") if multi_schema else None
            try:
                element_records = []
                active_view = doc.ActiveView
                if active_view:
                    collector = FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType()
                else:
                    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
                for elem in collector:
                    cat = elem.Category
                    if not cat or cat.CategoryType != CategoryType.Model:
                        continue
                    if not elem.IsValidObject:
                        continue
                    if is_camera_element(elem):
                        continue
                    entries = load_element_classifications(
                        elem,
                        legacy_schema=legacy_schema,
                        legacy_fields=legacy_fields,
                        multi_schema=multi_schema,
                        multi_items_field=multi_items_field
                    )
                    cat_name = (cat.Name if cat else "")
                    try:
                        elem_name = elem.Name or str(elem.Id.IntegerValue)
                    except Exception:
                        elem_name = str(elem.Id.IntegerValue)
                    element_records.append((elem, entries, cat_name, elem_name))

                element_records.sort(key=lambda rec: (
                    0 if rec[1] else 1,
                    (rec[2] or "").lower(),
                    (rec[3] or "").lower()
                ))

                items = []
                for elem, entries, _, _ in element_records:
                    if entries:
                        for idx, entry in enumerate(entries):
                            items.append(OverviewItem(elem, classification_item=entry, show_element_data=(idx == 0)))
                    else:
                        items.append(OverviewItem(elem, classification_item=None, show_element_data=True))

                for item in items:
                    self.OverviewList.Items.Add(item)

                classified = sum(1 for _, entries, _, _ in element_records if entries)
                total = len(element_records)
                total_cls_rows = sum(len(entries) for _, entries, _, _ in element_records)
                view_name = active_view.Name if active_view else "Active View"
                self.OverviewStats.Text = u"{} of {} model elements classified in {} ({} classifications)".format(
                    classified, total, view_name, total_cls_rows
                )
                self._set_status(u"Overview: {} classified, {} unclassified".format(
                    classified, total - classified))
            except Exception as ex:
                self._set_status(u"Overview error: {}".format(ex))

    last_dict_uri = ""

    while True:
        dlg = ClassificationDialog(preselected_dict_uri=last_dict_uri)
        dlg.ShowDialog()

        req = dlg._deferred_pick_request
        if not req:
            break

        last_dict_uri = (req.get("dict_data", {}) or {}).get("uri", "") or last_dict_uri

        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Select elements in the model and confirm"
            )
        except OperationCanceledException:
            continue
        except Exception:
            forms.alert("Element selection failed unexpectedly.", title="Error")
            continue

        if not refs:
            forms.alert("No elements selected.", title="Notice")
            continue

        picked_ids = [r.ElementId for r in refs if r and r.ElementId]
        if not picked_ids:
            forms.alert("No elements selected.", title="Notice")
            continue

        success, errors, tx_ex = apply_classification_to_elements(
            picked_ids,
            req.get("class_data", {}),
            req.get("dict_data", {})
        )
        if tx_ex is not None:
            forms.alert(
                "Classification write failed due to an unexpected Revit transaction error.",
                title="Error"
            )
            continue

        msg = "Classified {} element(s).".format(success)
        if errors:
            msg += " {} error(s).".format(errors)
        forms.alert(msg, title="Done")


main()