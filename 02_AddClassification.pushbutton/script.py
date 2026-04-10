# -*- coding: utf-8 -*-
__title__ = "Add\nClassification"
__doc__ = "Add external classification to selected elements"

from pyrevit import forms
import properties_editor as pe
import properties_editor_dialog as ped
import threading
import shared_text_utils as text_utils


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

    safe_unicode = text_utils.safe_unicode
    safe_query_value = text_utils.safe_query_value
    decode_escaped_text = text_utils.decode_escaped_text

    def safe_json_dumps(obj):
        def normalize(item):
            if item is None:
                return None
            if isinstance(item, (bool, int, long, float)):
                return item
            if isinstance(item, (unicode, str)):
                return safe_unicode(item)
            if isinstance(item, list):
                return [normalize(x) for x in item]
            if isinstance(item, tuple):
                return [normalize(x) for x in item]
            if isinstance(item, dict):
                result = {}
                for key, value in item.items():
                    result[safe_unicode(key)] = normalize(value)
                return result
            return safe_unicode(item)

        return _json.dumps(normalize(obj), ensure_ascii=True)

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
                                    "code": safe_unicode(p.get("code") or "").strip(),
                                    "name": safe_unicode(p.get("name") or "").strip(),
                                    "class_uri": safe_unicode(p.get("class_uri") or "").strip(),
                                    "dict_name": safe_unicode(p.get("dict_name") or "").strip(),
                                    "dict_uri": safe_unicode(p.get("dict_uri") or "").strip(),
                                    "parent_class_uri": safe_unicode(p.get("parent_class_uri") or "").strip(),
                                    "parent_class_code": safe_unicode(p.get("parent_class_code") or "").strip(),
                                    "parent_class_name": safe_unicode(p.get("parent_class_name") or "").strip(),
                                    "ancestor_classes": list(p.get("ancestor_classes") or []) if isinstance(p.get("ancestor_classes"), list) else [],
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
                            "code": safe_unicode(code),
                            "name": safe_unicode(ent.Get[System.String](legacy_fields["Name"]) or ""),
                            "class_uri": safe_unicode(ent.Get[System.String](legacy_fields["ClassUri"]) or ""),
                            "dict_name": safe_unicode(ent.Get[System.String](legacy_fields["DictionaryName"]) or ""),
                            "dict_uri": safe_unicode(ent.Get[System.String](legacy_fields["DictionaryUri"]) or ""),
                            "parent_class_uri": "",
                            "parent_class_code": "",
                            "parent_class_name": "",
                            "ancestor_classes": [],
                        })
        except Exception:
            pass

        return items

    def store_element_classifications(elem, items, legacy_schema=None, legacy_fields=None, multi_schema=None, multi_items_field=None):
        clean_items = []
        seen = set()
        for it in items:
            d_name = safe_unicode(it.get("dict_name") or "").strip()
            d_uri = safe_unicode(it.get("dict_uri") or "").strip()
            code = safe_unicode(it.get("code") or "").strip()
            name = safe_unicode(it.get("name") or "").strip()
            class_uri = safe_unicode(it.get("class_uri") or "").strip()
            parent_class_uri = safe_unicode(it.get("parent_class_uri") or "").strip()
            parent_class_code = safe_unicode(it.get("parent_class_code") or "").strip()
            parent_class_name = safe_unicode(it.get("parent_class_name") or "").strip()
            raw_ancestors = it.get("ancestor_classes") or []
            ancestor_classes = []
            if isinstance(raw_ancestors, list):
                for ancestor in raw_ancestors:
                    if not isinstance(ancestor, dict):
                        continue
                    ancestor_classes.append({
                        "uri": safe_unicode(ancestor.get("uri") or "").strip(),
                        "code": safe_unicode(ancestor.get("code") or "").strip(),
                        "name": safe_unicode(ancestor.get("name") or "").strip(),
                    })
            if not code or not d_name:
                continue
            key = (normalize_text(d_name), normalize_text(d_uri), normalize_text(code), normalize_text(name), normalize_text(class_uri))
            if key in seen:
                continue
            seen.add(key)
            clean_items.append({
                "code": code,
                "name": name,
                "class_uri": class_uri,
                "dict_name": d_name,
                "dict_uri": d_uri,
                "parent_class_uri": parent_class_uri,
                "parent_class_code": parent_class_code,
                "parent_class_name": parent_class_name,
                "ancestor_classes": ancestor_classes,
            })

        if multi_schema and multi_items_field:
            ment = Entity(multi_schema)
            ment.Set[System.String](multi_items_field, safe_json_dumps(clean_items))
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
        code      = safe_unicode((cls_data or {}).get("code", (cls_data or {}).get("referenceCode", "")))
        name      = safe_unicode((cls_data or {}).get("name", ""))
        class_uri = safe_unicode((cls_data or {}).get("classUri", (cls_data or {}).get("uri", "")))
        parent_class_uri = safe_unicode((cls_data or {}).get("_parentClassUri", (cls_data or {}).get("parentClassUri", "")))
        parent_class_code = safe_unicode((cls_data or {}).get("_parentClassCode", (cls_data or {}).get("parentClassCode", "")))
        parent_class_name = safe_unicode((cls_data or {}).get("_parentClassName", (cls_data or {}).get("parentClassName", "")))
        ancestor_classes = list((cls_data or {}).get("_ancestorClasses", []) or [])
        dict_name = safe_unicode(dict_name)
        dict_uri  = safe_unicode(dict_uri)

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

        def same_class(item):
            item_class_uri = normalize_text(item.get("class_uri", ""))
            if class_uri and item_class_uri:
                return item_class_uri == normalize_text(class_uri)
            return (
                normalize_text(item.get("dict_uri", "")) == target_dict_uri_norm and
                normalize_text(item.get("dict_name", "")) == target_dict_name_norm and
                normalize_text(item.get("code", "")) == normalize_text(code) and
                normalize_text(item.get("name", "")) == normalize_text(name)
            )

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
                    merged = [it for it in existing if not same_class(it)]
                    merged.append({
                        "code": code or "",
                        "name": name or "",
                        "class_uri": class_uri or "",
                        "dict_name": dict_name or "",
                        "dict_uri": dict_uri or "",
                        "parent_class_uri": parent_class_uri or "",
                        "parent_class_code": parent_class_code or "",
                        "parent_class_name": parent_class_name or "",
                        "ancestor_classes": ancestor_classes,
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
            depth = int(data.get("_depth", 0) or 0)
            raw_name = safe_unicode(data.get("name", "") or "")
            if depth > 0:
                self.name = u"{}↳ {}".format(u"   " * depth, raw_name)
            else:
                self.name = raw_name
            self.code            = safe_unicode(data.get("code", data.get("referenceCode", "")) or "")
            self.descriptionPart = safe_unicode(data.get("descriptionPart", "") or "")
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
        "classes_by_uri": {},
        "class_details_by_key": {},
        "no_properties_keys": set(),
        "prefetch_inflight_keys": set(),
    }

    def _extract_class_detail_candidate(payload_obj):
        if isinstance(payload_obj, dict):
            cps = payload_obj.get("classProperties", []) or []
            if isinstance(cps, list):
                return payload_obj

            classes = payload_obj.get("classes", []) or []
            if isinstance(classes, list):
                for c in classes:
                    if isinstance(c, dict) and isinstance(c.get("classProperties", []), list):
                        return c

        if isinstance(payload_obj, list):
            for c in payload_obj:
                if isinstance(c, dict) and isinstance(c.get("classProperties", []), list):
                    return c
        return None

    def _query_class_detail_payload(class_uri, dict_uri, timeout_sec=8):
        import urllib2 as _ul
        import urllib as _ulbase
        import json as _json

        endpoints = [
            url.rstrip("/") + "/api/Class/v1",
            url.rstrip("/") + "/api/Class/v1/",
        ]

        attempts = []
        p1 = {"Uri": class_uri, "IncludeClassProperties": "true"}
        if dict_uri:
            p1["DictionaryUri"] = dict_uri
        attempts.append(p1)
        attempts.append({"Uri": class_uri, "IncludeClassProperties": "true"})
        attempts.append({"ClassUri": class_uri, "IncludeClassProperties": "true"})

        best_payload = None
        best_count = 0
        saw_explicit_empty = False

        for endpoint in endpoints:
            for params in attempts:
                try:
                    safe = {}
                    for k, v in params.items():
                        safe[safe_query_value(k)] = safe_query_value(v)

                    full = endpoint + "?" + _ulbase.urlencode(safe)
                    resp = _ul.urlopen(full.encode("utf-8") if isinstance(full, unicode) else full, timeout=timeout_sec)
                    raw = resp.read()

                    payload = None
                    try:
                        payload = _json.loads(raw.decode("utf-8"))
                    except Exception:
                        try:
                            payload = _json.loads(raw)
                        except Exception:
                            payload = None

                    candidate = _extract_class_detail_candidate(payload)
                    if not isinstance(candidate, dict):
                        continue

                    count = len(list(candidate.get("classProperties", []) or []))
                    if count == 0:
                        saw_explicit_empty = True
                    if best_payload is None or count > best_count:
                        best_payload = candidate
                        best_count = count
                    if count > 0:
                        return best_payload, best_count, saw_explicit_empty
                except Exception:
                    pass

        return best_payload, best_count, saw_explicit_empty

    def _get_class_detail_cached(class_uri, dict_uri, allow_network=True, timeout_sec=8):
        cache_key = u"{}|{}".format(
            normalize_text(dict_uri or ""),
            normalize_text(class_uri or "")
        )

        details_cache = session_cache.setdefault("class_details_by_key", {})
        no_props_cache = session_cache.setdefault("no_properties_keys", set())

        if cache_key in no_props_cache:
            return "no-properties", None, cache_key

        cached_detail = details_cache.get(cache_key)
        if isinstance(cached_detail, dict):
            return "cached", cached_detail, cache_key

        if not allow_network:
            return "none", None, cache_key

        payload, count, saw_explicit_empty = _query_class_detail_payload(class_uri, dict_uri, timeout_sec=timeout_sec)
        if isinstance(payload, dict) and count > 0:
            details_cache[cache_key] = dict(payload)
            return "fetched", payload, cache_key

        if saw_explicit_empty:
            no_props_cache.add(cache_key)
            return "no-properties", None, cache_key

        return "none", None, cache_key

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
                    safe[safe_query_value(k)] = safe_query_value(v)
                full = full + "?" + _ulbase.urlencode(safe)
            resp = _ul.urlopen(full.encode("utf-8") if isinstance(full, unicode) else full, timeout=30)
            raw  = resp.read()
            return _json.loads(raw.decode("utf-8"))

        def _get(self, endpoint, params=None):
            return self._api_get(endpoint, params)

        def _set_status(self, msg):
            self.StatusText.Text = msg

        def _prefetch_selected_class_details(self, cls_data):
            try:
                class_uri = safe_unicode((cls_data or {}).get("uri", (cls_data or {}).get("classUri", ""))).strip()
                if not class_uri:
                    return

                idx = self.DictionaryComboBox.SelectedIndex
                dict_data = self._dicts[idx] if (idx >= 0 and idx < len(self._dicts)) else {}
                dict_uri = safe_unicode((dict_data or {}).get("uri", "")).strip()

                status, _, cache_key = _get_class_detail_cached(class_uri, dict_uri, allow_network=False)
                if status in ("cached", "no-properties"):
                    return

                inflight = session_cache.setdefault("prefetch_inflight_keys", set())
                if cache_key in inflight:
                    return
                inflight.add(cache_key)

                def _worker():
                    try:
                        _get_class_detail_cached(class_uri, dict_uri, allow_network=True, timeout_sec=6)
                    except Exception:
                        pass
                    finally:
                        try:
                            inflight.discard(cache_key)
                        except Exception:
                            pass

                th = threading.Thread(target=_worker)
                th.daemon = True
                th.start()
            except Exception:
                pass

        def _load_dictionaries(self):
            self._set_status("Loading dictionaries...")
            try:
                dicts = session_cache.get("dicts")
                if dicts is None:
                    result = self._get("/api/Dictionary/v1")
                    dicts = result.get("dictionaries", []) if isinstance(result, dict) else []
                    dicts = sorted(
                        dicts,
                        key=lambda d: safe_unicode((d.get("name") if isinstance(d, dict) else d) or u"").lower()
                    )
                    session_cache["dicts"] = dicts
                if not dicts:
                    self._set_status("No dictionaries found.")
                    return
                self._dicts = dicts
                self._suspend_dict_selection_handler = True
                self.DictionaryComboBox.Items.Clear()
                for d in self._dicts:
                    _dname = safe_unicode(d.get("name") or d)
                    _dver = safe_unicode(d.get("version") or d.get("releaseDate") or "").strip()
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
                    result = self._get("/api/Dictionary/v1/Classes", {"Uri": dict_uri, "UseNestedClasses": "true"})
                    root_classes = result.get("classes", []) if isinstance(result, dict) else []
                    classes = self._flatten_classes(root_classes)
                    if not any((c.get("_depth", 0) or 0) > 0 for c in classes):
                        classes.sort(key=lambda c: safe_unicode(c.get("name") or u"").lower())
                    classes_by_uri[dict_uri] = classes
                self._all_classes = classes
                self._populate_list(classes)
                self.SearchBox.IsEnabled = True
                has_nested = any((c.get("_depth", 0) or 0) > 0 for c in classes)
                if has_nested:
                    self._set_status("Loaded {} classes (nested hierarchy).".format(len(classes)))
                else:
                    self._set_status("Loaded {} classes.".format(len(classes)))
            except Exception as ex:
                self._set_status("Error: {}".format(ex))

        def _flatten_classes(self, classes, depth=0, parent=None, ancestors=None):
            flat = []
            ancestors = list(ancestors or [])
            for cls in list(classes or []):
                if not isinstance(cls, dict):
                    continue

                entry = dict(cls)
                entry["_depth"] = int(depth)
                entry["_parentClassUri"] = safe_unicode((parent or {}).get("uri") or "") if parent else ""
                entry["_parentClassCode"] = safe_unicode((parent or {}).get("code") or "") if parent else ""
                entry["_parentClassName"] = safe_unicode((parent or {}).get("name") or "") if parent else ""
                entry["_ancestorClasses"] = list(ancestors)
                flat.append(entry)

                node = {
                    "uri": safe_unicode(cls.get("uri") or cls.get("classUri") or ""),
                    "code": safe_unicode(cls.get("code") or cls.get("referenceCode") or ""),
                    "name": safe_unicode(cls.get("name") or ""),
                }
                next_ancestors = list(ancestors)
                if node.get("uri"):
                    next_ancestors.append(node)

                children = cls.get("children", []) or []
                if isinstance(children, list) and children:
                    flat.extend(self._flatten_classes(children, depth + 1, cls, next_ancestors))

            return flat

        def _populate_list(self, classes):
            self.ClassListBox.Items.Clear()
            for c in classes:
                self.ClassListBox.Items.Add(ClassItem(c))

        def _get_selected_class(self):
            item = self.ClassListBox.SelectedItem
            return item.data if item else None

        def _apply_to_elements(self, element_ids, show_done_alert=True, open_properties_after=False):
            cls_data = self._get_selected_class()
            if not cls_data:
                forms.alert("Please select a class first.", title="Notice")
                return 0, 0, None
            idx       = self.DictionaryComboBox.SelectedIndex
            dict_data = self._dicts[idx] if (idx >= 0 and idx < len(self._dicts)) else {}

            success, errors, tx_ex = apply_classification_to_elements(element_ids, cls_data, dict_data)
            if tx_ex is not None:
                forms.alert(
                    "Classification write failed due to an unexpected Revit transaction error.",
                    title="Error"
                )
                return 0, 0, tx_ex
            msg = "Classified {} element(s).".format(success)
            if errors:
                msg += " {} error(s).".format(errors)
            self._set_status(msg)
            if show_done_alert:
                forms.alert(msg, title="Done")

            if open_properties_after and success > 0:
                run_post_classification_property_flow(
                    list(element_ids or []),
                    cls_data,
                    dict_data,
                    success
                )

            return success, errors, None

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
                        removed_class_uris = []
                        for it in existing:
                            matched = False
                            for sel in selectors:
                                if _matches_selector(it, sel):
                                    matched = True
                                    break
                            if matched:
                                local_removed += 1
                                try:
                                    rem_uri = (it.get("class_uri", "") or "").strip()
                                    if rem_uri:
                                        removed_class_uris.append(rem_uri)
                                except Exception:
                                    pass
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

                        # Keep properties storage consistent with removed classifications.
                        for rem_uri in set(removed_class_uris):
                            try:
                                pe.delete_class_properties(elem, rem_uri)
                            except Exception:
                                pass

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
            query = safe_unicode(self.SearchBox.Text or "").strip().lower()
            if not query:
                self._populate_list(self._all_classes)
                return
            filtered = [
                c for c in self._all_classes
                if query in safe_unicode(c.get("name", "") or "").lower()
                or query in safe_unicode(c.get("code", "") or "").lower()
                or query in safe_unicode(c.get("referenceCode", "") or "").lower()
            ]
            self._populate_list(filtered)

        def ClassListBox_SelectionChanged(self, sender, args):
            item = self.ClassListBox.SelectedItem
            can  = item is not None
            self.AssignToSelectionBtn.IsEnabled = can
            self.SelectAndAssignBtn.IsEnabled   = can
            if item:
                d = item.data
                self.DetailCode.Text = safe_unicode(d.get("code", d.get("referenceCode", "")))
                self.DetailName.Text = safe_unicode(d.get("name", ""))
                self.DetailUri.Text  = safe_unicode(d.get("classUri", d.get("uri", "")))
                self.ClassDetailPanel.Visibility = System.Windows.Visibility.Visible
                self._prefetch_selected_class_details(d)
            else:
                self.ClassDetailPanel.Visibility = System.Windows.Visibility.Collapsed

        def _should_open_properties(self):
            try:
                return bool(self.AutoOpenPropertiesCheckBox.IsChecked)
            except Exception:
                return True

        def AssignToSelection_Click(self, sender, args):
            sel_ids = list(uidoc.Selection.GetElementIds())
            if not sel_ids:
                forms.alert("No elements selected.\nPlease select elements in Revit first.", title="Notice")
                return
            self._apply_to_elements(sel_ids, open_properties_after=self._should_open_properties())

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
                },
                "open_properties_after": self._should_open_properties(),
            }
            self.Close()

        def Close_Click(self, sender, args):
            self.Close()

        def Tab_SelectionChanged(self, sender, args):
            if args.OriginalSource != self.MainTabControl:
                return
            if self.MainTabControl.SelectedIndex == 1:
                self._load_overview()
            elif self.MainTabControl.SelectedIndex == 2:
                self._load_properties()

        def RefreshOverview_Click(self, sender, args):
            self._load_overview()

        def RefreshProperties_Click(self, sender, args):
            self._load_properties()

        def _load_properties(self):
            import clr
            clr.AddReference('PresentationFramework')
            clr.AddReference('PresentationCore')
            clr.AddReference('WindowsBase')
            from System.Windows import Thickness, FontWeights
            from System.Windows.Controls import TreeViewItem, Grid, ColumnDefinition, TextBlock
            from System.Windows import GridLength, GridUnitType
            from System.Windows.Media import FontFamily
            import Autodesk.Revit.DB as DB_inner

            self.PropertiesTree.Items.Clear()
            properties_count = 0

            def _build_property_row(name_text, value_text, type_text, is_header=False):
                row_grid = Grid()
                row_grid.MinWidth = 560
                col_name = ColumnDefinition()
                col_name.Width = GridLength(240, GridUnitType.Pixel)
                row_grid.ColumnDefinitions.Add(col_name)

                col_value = ColumnDefinition()
                col_value.Width = GridLength(1, GridUnitType.Star)
                row_grid.ColumnDefinitions.Add(col_value)

                col_type = ColumnDefinition()
                col_type.Width = GridLength(140, GridUnitType.Pixel)
                row_grid.ColumnDefinitions.Add(col_type)

                cell_name = TextBlock(Text=name_text)
                cell_name.Margin = Thickness(0, 0, 8, 0)
                cell_name.FontWeight = FontWeights.SemiBold if is_header else FontWeights.Normal
                cell_name.FontFamily = FontFamily("Consolas")
                Grid.SetColumn(cell_name, 0)

                cell_value = TextBlock(Text=value_text)
                cell_value.Margin = Thickness(0, 0, 8, 0)
                cell_value.FontWeight = FontWeights.SemiBold if is_header else FontWeights.Normal
                cell_value.FontFamily = FontFamily("Consolas")
                Grid.SetColumn(cell_value, 1)

                cell_type = TextBlock(Text=type_text)
                cell_type.FontWeight = FontWeights.SemiBold if is_header else FontWeights.Normal
                cell_type.FontFamily = FontFamily("Consolas")
                Grid.SetColumn(cell_type, 2)

                row_grid.Children.Add(cell_name)
                row_grid.Children.Add(cell_value)
                row_grid.Children.Add(cell_type)
                return row_grid

            def _decode_display_text(raw_value):
                txt = decode_escaped_text(raw_value)
                return txt if txt else u"(empty)"

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
                all_ids = list(DB_inner.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElementIds())

                for eid in all_ids:
                    try:
                        elem = doc.GetElement(eid)
                        if not elem or not elem.IsValidObject:
                            continue

                        elem_props = pe.load_all_class_properties(elem)
                        if not elem_props:
                            continue

                        class_name_by_uri = {}
                        try:
                            cls_entries = load_element_classifications(
                                elem,
                                legacy_schema=legacy_schema,
                                legacy_fields=legacy_fields,
                                multi_schema=multi_schema,
                                multi_items_field=multi_items_field
                            )
                            for cls_entry in cls_entries:
                                cls_uri = (cls_entry.get("class_uri") or "").strip()
                                cls_name = (cls_entry.get("name") or cls_entry.get("code") or "").strip()
                                if cls_uri and cls_name:
                                    class_name_by_uri[cls_uri] = cls_name
                        except Exception:
                            pass

                        try:
                            elem_name = elem.Name or "Unknown"
                        except Exception:
                            elem_name = "Unknown"

                        elem_item = TreeViewItem()
                        elem_item.Header = u"{}  [ID: {}]".format(elem_name, eid.IntegerValue)
                        elem_item.Tag = (eid.IntegerValue, "element")
                        elem_item.IsExpanded = True

                        for class_uri, property_data in elem_props.items():
                            class_item = TreeViewItem()
                            class_name = class_name_by_uri.get(class_uri, "")
                            if class_name:
                                class_item.Header = u"Class: {}".format(class_name)
                            else:
                                class_item.Header = u"Class: {}".format(class_uri[:50] + "..." if len(class_uri) > 50 else class_uri)
                            class_item.ToolTip = class_uri
                            class_item.Tag = (eid.IntegerValue, class_uri, "class")
                            class_item.IsExpanded = True

                            psets = property_data.get("property_sets", [])
                            for pset in psets:
                                pset_item = TreeViewItem()
                                enabled_count = len([p for p in pset.get("properties", []) if p.get("enabled")])
                                pset_item.Header = u"{}  ({} enabled)".format(pset.get("name", "PropertySet"), enabled_count)
                                pset_item.Tag = (eid.IntegerValue, class_uri, pset.get("name", ""), "pset")
                                pset_item.IsExpanded = True

                                pset_header_item = TreeViewItem()
                                pset_header_item.Header = _build_property_row("Name", "Value", "Type", is_header=True)
                                pset_header_item.IsEnabled = False
                                pset_header_item.Focusable = False
                                pset_item.Items.Add(pset_header_item)

                                for prop in pset.get("properties", []):
                                    has_value = prop.get("value") not in (None, "")
                                    if prop.get("enabled") or has_value:
                                        prop_item = TreeViewItem()
                                        prop_name = prop.get("name", "")
                                        prop_val = prop.get("value", "")
                                        prop_type = prop.get("dataType", "")
                                        if not prop_type or str(prop_type).strip().lower() in ("none", "null", ""):
                                            prop_type = "undefined"

                                        prop_val_display = _decode_display_text(prop_val)

                                        prop_item.Header = _build_property_row(
                                            unicode(prop_name or ""),
                                            unicode(prop_val_display or ""),
                                            unicode(prop_type or "undefined")
                                        )
                                        prop_item.Tag = (eid.IntegerValue, class_uri, pset.get("name", ""), prop_name, prop_val)
                                        pset_item.Items.Add(prop_item)
                                        properties_count += 1

                                class_item.Items.Add(pset_item)

                            elem_item.Items.Add(class_item)

                        self.PropertiesTree.Items.Add(elem_item)
                    except Exception:
                        pass

                if properties_count > 0:
                    self.PropertiesStatusText.Text = u"Showing: {} configured properties on {} element(s)".format(
                        properties_count, self.PropertiesTree.Items.Count
                    )
                else:
                    self.PropertiesStatusText.Text = "No properties configured yet. Configure in classification dialog."
            except Exception as ex:
                self.PropertiesStatusText.Text = "Error loading properties: {}".format(str(ex)[:80])

        def PropertiesTree_MouseDoubleClick(self, sender, args):
            import Autodesk.Revit.DB as DB_local
            from Autodesk.Revit.DB import Transaction as TX_local

            def _decode_for_edit(raw_value):
                return decode_escaped_text(raw_value)

            selected = sender.SelectedItem
            if not selected or not hasattr(selected, 'Tag') or not selected.Tag:
                return

            tag = selected.Tag
            if not isinstance(tag, tuple) or len(tag) != 5:
                return

            try:
                eid_int, class_uri, pset_name, prop_name, current_val = tag
                eid = DB_local.ElementId(eid_int)
                elem = doc.GetElement(eid)
                if not elem or not elem.IsValidObject:
                    return

                current_display = _decode_for_edit(current_val)

                from pyrevit import forms as prf
                new_value = prf.ask_for_string(
                    current_display or "",
                    title="Edit Property Value",
                    prompt=u"New value for '{}':\n(Current: {})".format(prop_name, current_display or "(empty)")
                )

                if new_value is not None:
                    tx = TX_local(doc, "Update Property Value")
                    tx.Start()
                    ok = False
                    try:
                        ok = bool(pe.update_property_value(elem, class_uri, pset_name, prop_name, new_value))
                        if ok:
                            tx.Commit()
                        else:
                            tx.RollBack()
                    except Exception:
                        try:
                            tx.RollBack()
                        except Exception:
                            pass
                        raise

                    if ok:
                        self._load_properties()
                        self.PropertiesStatusText.Text = u"Updated: {}".format(prop_name)
                    else:
                        self.PropertiesStatusText.Text = u"Update failed: {}".format(prop_name)
            except Exception:
                self.PropertiesStatusText.Text = u"Error: Operation failed"

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

    def run_post_classification_property_flow(element_ids, class_data, dict_data, success_count):
        """Open property dialog after successful classification assignment."""
        try:
            if not element_ids or success_count <= 0:
                return

            class_uri = (class_data or {}).get("uri", "") or (class_data or {}).get("classUri", "")
            dict_uri = (dict_data or {}).get("uri", "")
            if not class_uri:
                return

            dialog_class_data = dict(class_data or {})
            status, payload_detail, _ = _get_class_detail_cached(
                class_uri,
                dict_uri,
                allow_network=True,
                timeout_sec=8
            )

            if status == "no-properties":
                forms.alert(
                    "The selected class has no properties.",
                    title="Properties"
                )
                return

            if isinstance(payload_detail, dict):
                merged_cached = dict(dialog_class_data or {})
                merged_cached.update(payload_detail)
                merged_cached["classProperties"] = list(payload_detail.get("classProperties", []) or [])
                dialog_class_data = merged_cached

            if "uri" not in dialog_class_data:
                dialog_class_data["uri"] = class_uri

            current_props = list((dialog_class_data or {}).get("classProperties", []) or [])
            if "classProperties" in (dialog_class_data or {}) and not current_props:
                forms.alert(
                    "The selected class has no properties.",
                    title="Properties"
                )
                return

            prop_dialog = ped.PropertyEditorDialog(doc, element_ids, dialog_class_data, dict_data or {})
            prop_dialog.ShowDialog()
        except Exception as ex:
            forms.alert("Classification saved, but the property dialog could not be opened.\n{}".format(ex), title="Properties")

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

        if req.get("open_properties_after", True):
            run_post_classification_property_flow(
                picked_ids,
                req.get("class_data", {}),
                req.get("dict_data", {}),
                success
            )


main()