# -*- coding: utf-8 -*-
__title__ = "Export\nIFC"
__doc__ = "Exports IFC with external classification."

import os
from pyrevit import forms


def main():
    import re
    import uuid
    import datetime
    import json
    import clr
    clr.AddReference("System")
    clr.AddReference("System.Windows.Forms")
    import System
    from System.Windows.Forms import SaveFileDialog, DialogResult
    clr.AddReference("RevitAPI")

    import Autodesk.Revit.DB as DB
    from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, Entity, AccessLevel
    try:
        from Autodesk.Revit.DB.ExtensibleStorage import DataStorage
    except ImportError:
        DataStorage = DB.DataStorage

    try:
        text_type = unicode
    except NameError:
        text_type = str

    doc = __revit__.ActiveUIDocument.Document

    url_schema_guid = "12345678-1234-1234-1234-1234567890ab"
    url_field_name = "APIUrl"
    cls_schema_guid = "ABCDEF12-3456-7890-ABCD-EF1234567890"
    cls_multi_schema_guid = "ABCDEF12-3456-7890-ABCD-EF1234567891"
    ifc_cls_schema_guid = "9A5A28C2-DDAC-4828-8B8A-3EE97118017A"
    ifc_builtin_sp = r"C:\Program Files\Autodesk\Revit 2025\IFC Shared Parameters-RevitIFCBuiltIn_ALL.txt"

    def get_ext_def(shared_param_file, group_name, pname, create_if_missing):
        app = doc.Application
        old = app.SharedParametersFilename
        try:
            app.SharedParametersFilename = shared_param_file
            df = app.OpenSharedParameterFile()
            if df is None:
                return None
            if group_name:
                grp = df.Groups.get_Item(group_name) or df.Groups.Create(group_name)
                defn = grp.Definitions.get_Item(pname)
                if not defn and create_if_missing:
                    opt = DB.ExternalDefinitionCreationOptions(pname, DB.SpecTypeId.String.Text)
                    defn = grp.Definitions.Create(opt)
                return defn
            for grp in df.Groups:
                defn = grp.Definitions.get_Item(pname)
                if defn:
                    return defn
            return None
        finally:
            app.SharedParametersFilename = old

    def get_stored_api_url():
        try:
            schema = Schema.Lookup(System.Guid(url_schema_guid))
            if schema:
                ent = doc.ProjectInformation.GetEntity(schema)
                if ent and ent.IsValid():
                    field = schema.GetField(url_field_name)
                    if field:
                        return ent.Get[System.String](field)
        except Exception:
            pass
        return None

    def api_get_json(base_url, endpoint, params=None):
        import urllib2 as _ul
        import urllib as _ub
        import json as _json
        full = base_url.rstrip("/") + endpoint
        if params:
            safe = {}
            for k, v in params.items():
                if isinstance(k, text_type):
                    k = k.encode("utf-8")
                if isinstance(v, text_type):
                    v = v.encode("utf-8")
                safe[k] = v
            full = full + "?" + _ub.urlencode(safe)
        resp = _ul.urlopen(full.encode("utf-8") if isinstance(full, text_type) else full, timeout=30)
        raw = resp.read()
        return _json.loads(raw.decode("utf-8"))

    def first_nonempty(data, keys):
        if not isinstance(data, dict):
            return None
        for key in keys:
            val = data.get(key)
            if val is None:
                continue
            sval = text_type(val).strip() if not isinstance(val, text_type) else val.strip()
            if sval:
                return sval
        return None

    def parse_date_value(raw_date):
        if not raw_date:
            return None
        try:
            s = text_type(raw_date).strip()
            if len(s) >= 10 and re.match(r"^\d{4}-\d{2}-\d{2}", s):
                d = datetime.datetime.strptime(s[:10], "%Y-%m-%d").date()
                return (d.day, d.month, d.year, s[:10])
        except Exception:
            pass
        return None

    def get_dictionary_metadata(dict_name, dict_uri):
        api_url = get_stored_api_url()
        if not api_url:
            return (None, None, None)
        try:
            payload = api_get_json(api_url, "/api/Dictionary/v1")
            dictionaries = payload.get("dictionaries", []) if isinstance(payload, dict) else []
            selected_dict = None
            for d in dictionaries:
                uri_val = (d.get("uri") or "").strip()
                name_val = (d.get("name") or "").strip()
                if dict_uri and uri_val and uri_val == dict_uri:
                    selected_dict = d
                    break
                if dict_name and name_val and name_val == dict_name:
                    selected_dict = d
                    break
            if not selected_dict:
                return (None, None, None)

            edition = first_nonempty(selected_dict, [
                "version", "edition", "releaseVersion", "revision", "dictionaryVersion"
            ])
            raw_date = first_nonempty(selected_dict, [
                "versionDate", "editionDate", "releaseDate", "publishedOn", "publicationDate", "lastUpdated", "updatedAt", "date"
            ])
            parsed_date = parse_date_value(raw_date)
            date_parts = None
            date_text = None
            if parsed_date:
                date_parts = (parsed_date[0], parsed_date[1], parsed_date[2])
                date_text = parsed_date[3]

            return (edition, date_parts, date_text)
        except Exception:
            return (None, None, None)

    def ensure_bound(cat, pname, shared_param_file, group_name=None, create_if_missing=False):
        it = doc.ParameterBindings.ForwardIterator()
        while it.MoveNext():
            if it.Key.Name == pname:
                binding = it.Current
                if not binding.Categories.Contains(cat):
                    binding.Categories.Insert(cat)
                    doc.ParameterBindings.ReInsert(it.Key, binding)
                return
        defn = get_ext_def(shared_param_file, group_name, pname, create_if_missing)
        if defn:
            cats = doc.Application.Create.NewCategorySet()
            cats.Insert(cat)
            doc.ParameterBindings.Insert(defn, doc.Application.Create.NewInstanceBinding(cats))

    def get_or_create_ifc_cls_schema():
        schema = Schema.Lookup(System.Guid(ifc_cls_schema_guid))
        if schema:
            return schema
        sb = SchemaBuilder(System.Guid(ifc_cls_schema_guid))
        sb.SetSchemaName("IFCClassification")
        sb.SetReadAccessLevel(AccessLevel.Public)
        sb.SetWriteAccessLevel(AccessLevel.Public)
        sb.AddSimpleField("ClassificationName", System.String)
        sb.AddSimpleField("ClassificationSource", System.String)
        sb.AddSimpleField("ClassificationEdition", System.String)
        sb.AddSimpleField("ClassificationEditionDate_Day", System.Int32)
        sb.AddSimpleField("ClassificationEditionDate_Month", System.Int32)
        sb.AddSimpleField("ClassificationEditionDate_Year", System.Int32)
        sb.AddSimpleField("ClassificationLocation", System.String)
        sb.AddSimpleField("ClassificationFieldName", System.String)
        return sb.Finish()

    def write_ifc_classification_storage(all_systems_data):
        # Creates one IFCClassification DataStorage entry per classification system.
        # The IFC exporter maps each entry to an IfcClassification via ClassificationFieldName.
        schema = get_or_create_ifc_cls_schema()
        data_ids = list(DB.FilteredElementCollector(doc).OfClass(DataStorage).ToElementIds())
        to_delete = []
        for ds_id in data_ids:
            try:
                ds = doc.GetElement(ds_id)
                if not ds:
                    continue
                ent = ds.GetEntity(schema)
                if ent and ent.IsValid():
                    to_delete.append(ds.Id)
            except Exception:
                pass
        if to_delete:
            ids_to_delete = System.Collections.Generic.List[DB.ElementId]()
            for eid in to_delete:
                ids_to_delete.Add(eid)
            doc.Delete(ids_to_delete)
        for sd in all_systems_data:
            d_name = sd.get("dict_name") or ""
            d_uri = sd.get("dict_uri") or ""
            edition = sd.get("edition")
            edition_date_parts = sd.get("edition_date_parts")
            param_name = sd.get("param_name") or "ClassificationCode"
            ds = DataStorage.Create(doc)
            ent = Entity(schema)
            ent.Set[System.String](schema.GetField("ClassificationName"), d_name)
            ent.Set[System.String](schema.GetField("ClassificationSource"), "")
            if edition:
                ent.Set[System.String](schema.GetField("ClassificationEdition"), edition)
            if edition_date_parts and len(edition_date_parts) == 3:
                ent.Set[System.Int32](schema.GetField("ClassificationEditionDate_Day"), int(edition_date_parts[0]))
                ent.Set[System.Int32](schema.GetField("ClassificationEditionDate_Month"), int(edition_date_parts[1]))
                ent.Set[System.Int32](schema.GetField("ClassificationEditionDate_Year"), int(edition_date_parts[2]))
            ent.Set[System.String](schema.GetField("ClassificationLocation"), d_uri)
            ent.Set[System.String](schema.GetField("ClassificationFieldName"), param_name)
            ds.SetEntity(ent)

    def ifc_literal(value):
        if value is None:
            return "$"
        txt = text_type(value).replace("'", "''").strip()
        if not txt:
            return "$"
        return "'{}'".format(txt)

    def ifc_unquote(token):
        tok = (token or "").strip()
        if tok == "$":
            return ""
        if len(tok) >= 2 and tok[0] == "'" and tok[-1] == "'":
            return tok[1:-1].replace("''", "'")
        return tok

    def split_ifc_args(arg_text):
        args = []
        buf = []
        in_string = False
        paren_depth = 0
        i = 0
        while i < len(arg_text):
            ch = arg_text[i]
            if ch == "'":
                buf.append(ch)
                if in_string:
                    if i + 1 < len(arg_text) and arg_text[i + 1] == "'":
                        buf.append("'")
                        i += 1
                    else:
                        in_string = False
                else:
                    in_string = True
            elif not in_string and ch == "(":
                paren_depth += 1
                buf.append(ch)
            elif not in_string and ch == ")":
                if paren_depth > 0:
                    paren_depth -= 1
                buf.append(ch)
            elif not in_string and ch == "," and paren_depth == 0:
                args.append("".join(buf).strip())
                buf = []
            else:
                buf.append(ch)
            i += 1
        args.append("".join(buf).strip())
        return args

    def normalize_text(value):
        return (text_type(value or "").strip()).lower()

    def read_element_classifications(elem, legacy_schema=None, legacy_fields=None, multi_schema=None, multi_items_field=None):
        items = []

        try:
            if multi_schema and multi_items_field:
                ment = elem.GetEntity(multi_schema)
                if ment and ment.IsValid():
                    raw = ment.Get[System.String](multi_items_field) or ""
                    if raw.strip():
                        payload = json.loads(raw)
                        if isinstance(payload, list):
                            for p in payload:
                                if not isinstance(p, dict):
                                    continue
                                items.append({
                                    "dict_name": (p.get("dict_name") or "").strip(),
                                    "dict_uri": (p.get("dict_uri") or "").strip(),
                                    "code": (p.get("code") or "").strip(),
                                    "name": (p.get("name") or "").strip(),
                                    "class_uri": (p.get("class_uri") or "").strip(),
                                })
        except Exception:
            items = []

        if items:
            return items

        try:
            if legacy_schema and legacy_fields:
                ent = elem.GetEntity(legacy_schema)
                if ent and ent.IsValid():
                    code_val = ent.Get[System.String](legacy_fields["Code"]) or ""
                    if code_val:
                        items.append({
                            "dict_name": ent.Get[System.String](legacy_fields["DictionaryName"]) or "",
                            "dict_uri": ent.Get[System.String](legacy_fields["DictionaryUri"]) or "",
                            "code": code_val,
                            "name": ent.Get[System.String](legacy_fields["Name"]) or "",
                            "class_uri": ent.Get[System.String](legacy_fields["ClassUri"]) or "",
                        })
        except Exception:
            pass

        return items

    def new_ifc_guid_22():
        alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$"
        num = uuid.uuid4().int
        out = []
        for _ in range(22):
            out.append(alphabet[num & 63])
            num >>= 6
        return "".join(reversed(out))

    def is_ifc_guid_token(token):
        txt = ifc_unquote(token)
        if len(txt) != 22:
            return False
        for ch in txt:
            if not (ch.isdigit() or ('A' <= ch <= 'Z') or ('a' <= ch <= 'z') or ch in "_$"):
                return False
        return True

    def collect_class_uri_maps_from_elements(dict_name, dict_uri, element_ids):
        """
        Extract ClassUri mappings directly from ExtensibleStorage of elements,
        instead of querying the API. This is more robust because:
        1. URIs are already saved when elements were classified
        2. No API calls needed (no network/timeout issues)
        3. Exactly the URIs that were used during classification are re-exported
        """
        by_code = {}
        by_pair = {}
        
        _CLS_SCHEMA_GUID_STR = "ABCDEF12-3456-7890-ABCD-EF1234567890"
        schema = Schema.Lookup(System.Guid(_CLS_SCHEMA_GUID_STR))
        if not schema:
            return by_code, by_pair
        
        code_field = schema.GetField("Code")
        name_field = schema.GetField("Name")
        class_uri_field = schema.GetField("ClassUri")
        dict_name_field = schema.GetField("DictionaryName")
        dict_uri_field = schema.GetField("DictionaryUri")
        
        if not (code_field and class_uri_field):
            return by_code, by_pair
        
        for elem_id in element_ids:
            try:
                elem = doc.GetElement(elem_id)
                if not elem or not elem.IsValidObject:
                    continue
                entity = elem.GetEntity(schema)
                if not entity or not entity.IsValid():
                    continue
                
                # Filter by dict_name and dict_uri to get URIs for THIS system only
                stored_dict_name = entity.Get[System.String](dict_name_field) if dict_name_field else ""
                stored_dict_uri = entity.Get[System.String](dict_uri_field) if dict_uri_field else ""
                
                # Match this element's system to the target system
                if normalize_text(stored_dict_name) != normalize_text(dict_name):
                    continue
                if normalize_text(stored_dict_uri) != normalize_text(dict_uri):
                    continue
                
                # Extract the URI mapping from this element
                code_val = entity.Get[System.String](code_field) if code_field else ""
                name_val = entity.Get[System.String](name_field) if name_field else ""
                class_uri_val = entity.Get[System.String](class_uri_field) if class_uri_field else ""
                
                code_val = (code_val or "").strip()
                name_val = (name_val or "").strip()
                class_uri_val = (class_uri_val or "").strip()
                
                if not code_val or not class_uri_val:
                    continue
                
                code_key = normalize_text(code_val)
                name_key = normalize_text(name_val)
                
                by_code[code_key] = class_uri_val
                if name_key:
                    by_pair[(code_key, name_key)] = class_uri_val
            
            except Exception:
                pass
        
        return by_code, by_pair

    def collect_class_uri_maps_from_api(dict_name, dict_uri):
        by_code = {}
        by_pair = {}
        api_url = get_stored_api_url()
        if not api_url:
            return by_code, by_pair
        if not dict_uri and not dict_name:
            return by_code, by_pair

        def add_classes_from_payload(payload):
            classes = payload.get("classes", []) if isinstance(payload, dict) else []
            for c in classes:
                code_val = (c.get("code") or c.get("referenceCode") or "").strip()
                name_val = (c.get("name") or "").strip()
                class_uri_val = (c.get("classUri") or c.get("uri") or "").strip()
                if not code_val or not class_uri_val:
                    continue
                code_key = normalize_text(code_val)
                name_key = normalize_text(name_val)
                by_code[code_key] = class_uri_val
                if name_key:
                    by_pair[(code_key, name_key)] = class_uri_val

        if dict_uri:
            try:
                add_classes_from_payload(api_get_json(api_url, "/api/Dictionary/v1/Classes", {"Uri": dict_uri}))
            except Exception:
                pass
        if not by_code and dict_name:
            try:
                add_classes_from_payload(api_get_json(api_url, "/api/Dictionary/v1/Classes", {"Name": dict_name}))
            except Exception:
                pass

        return by_code, by_pair

    def _choose_ifc_export_options():
        """
        Tries to find the IFC Exporter add-in (IFCExporterUIOverride.dll, already
        loaded in Revit's process) and shows a SelectFromList dialog with all
        available export configurations (built-in + user-saved profiles).
        Applies the chosen profile to a fresh IFCExportOptions instance via
        IFCExportConfiguration.UpdateOptions().

        Falls back to basic IFC 4 options when the add-in is unavailable.
        Returns None when the user cancels the profile selection.
        """
        import System

        # Collect all loaded assemblies once.
        all_assemblies = list(System.AppDomain.CurrentDomain.GetAssemblies())
        all_asm_names = [asm.GetName().Name for asm in all_assemblies]

        # Search all IFC assemblies for the configuration map class.
        # Revit <=2024 add-in: IFCExportConfigurationMaps  (suffix "Maps")
        # Revit 2025 built-in: IFCExportConfigurationsMap  (prefix "Configurations")
        config_maps_type = None
        for asm, asm_name in zip(all_assemblies, all_asm_names):
            if "IFC" not in asm_name.upper():
                continue
            try:
                for t in asm.GetTypes():
                    if t.Name in ("IFCExportConfigurationMaps", "IFCExportConfigurationsMap"):
                        config_maps_type = t
                        break
            except Exception:
                pass
            if config_maps_type:
                break

        base_opts = DB.IFCExportOptions()

        if config_maps_type is None:
            # IFC Exporter not found in AppDomain – warn and fall back to IFC 4.
            forms.alert(
                u"IFC Exporter configuration not found.\n\n"
                u"The IFC export will use the built-in default (IFC 4).\n"
                u"Make sure the IFC Exporter is active in Revit.",
                title="Export IFC",
                warn_icon=True
            )
            try:
                base_opts.FileVersion = DB.IFCVersion.IFC4
            except Exception:
                pass
            return base_opts

        try:
            config_maps = config_maps_type()

            # Populate with built-in and user-saved profiles.
            try:
                config_maps_type.GetMethod("AddBuiltInConfigurations").Invoke(config_maps, None)
            except Exception:
                pass
            try:
                m = config_maps_type.GetMethod("AddSavedConfigurations")
                if m:
                    params = m.GetParameters()
                    if len(list(params)) == 1:
                        m.Invoke(config_maps, [doc])
                    else:
                        m.Invoke(config_maps, None)
            except Exception:
                pass

            configs = list(config_maps.Values)
            name_to_config = {str(c.Name): c for c in configs}
            config_names = sorted(name_to_config.keys())
        except Exception as e:
            forms.alert(
                u"Could not read IFC Exporter profiles:\n{}\n\n"
                u"Falling back to IFC 4 defaults.".format(e),
                title="Export IFC",
                warn_icon=True
            )
            try:
                base_opts.FileVersion = DB.IFCVersion.IFC4
            except Exception:
                pass
            return base_opts

        if not config_names:
            return base_opts

        selected_name = forms.SelectFromList.show(
            config_names,
            multiselect=False,
            title="IFC Export Configuration",
            button_name="Export"
        )
        if selected_name is None:
            return None  # user cancelled

        try:
            config = name_to_config[selected_name]
            config.UpdateOptions(base_opts, doc.ActiveView.Id)
        except Exception as e:
            forms.alert(
                u"Could not apply profile '{}':\n{}\n\n"
                u"Exporting with default IFC options.".format(selected_name, e),
                title="Export IFC",
                warn_icon=True
            )

        return base_opts

    def choose_ifc_output_path(default_doc_title):
        dialog = SaveFileDialog()
        dialog.Title = "Export IFC"
        dialog.Filter = "IFC files (*.ifc)|*.ifc"
        dialog.DefaultExt = "ifc"
        dialog.AddExtension = True
        dialog.OverwritePrompt = True
        dialog.FileName = u"{}_classified.ifc".format(default_doc_title or "Export")
        result = dialog.ShowDialog()
        if result != DialogResult.OK:
            return None
        file_path = dialog.FileName or ""
        if not file_path.lower().endswith(".ifc"):
            file_path += ".ifc"
        return file_path

    def postprocess_ifc_output(ifc_path, all_systems_data):
        if not os.path.exists(ifc_path):
            return (False, "IFC file not found")

        with open(ifc_path, "r") as f:
            content = f.read()

        updated = content

        has_cls_metadata = any(sd.get("edition") or sd.get("edition_date_text") for sd in all_systems_data)
        has_ref_uri_maps = any(sd.get("class_uri_by_code") or sd.get("class_uri_by_pair") for sd in all_systems_data)

        # Build map: IfcClassification entity-id (without #) → system_data, by matching name/URI.
        cls_line_regex = re.compile(
            r"(^\s*#(\d+)\s*=\s*IFCCLASSIFICATION\s*\()(.*?)(\)\s*;)",
            re.IGNORECASE | re.MULTILINE
        )
        cls_id_to_system = {}
        for cm in cls_line_regex.finditer(updated):
            cls_entity_id = cm.group(2)
            cls_args = split_ifc_args(cm.group(3))
            cls_name_in_file = ifc_unquote(cls_args[3]) if len(cls_args) > 3 else ""
            cls_loc_in_file  = ifc_unquote(cls_args[5]) if len(cls_args) > 5 else ""
            for sd in all_systems_data:
                d_name = sd.get("dict_name", "")
                d_uri  = sd.get("dict_uri", "")
                if (d_uri  and normalize_text(cls_loc_in_file)  == normalize_text(d_uri)) or \
                   (d_name and normalize_text(cls_name_in_file) == normalize_text(d_name)):
                    cls_id_to_system[cls_entity_id] = sd
                    break

        # 1) Update edition/date for matched IfcClassification entries in one pass.
        cls_subs = 0
        if has_cls_metadata and cls_id_to_system:
            def repl_cls(m):
                cls_entity_id = m.group(2)
                sd = cls_id_to_system.get(cls_entity_id)
                if sd is None:
                    return m.group(0)
                edition_value = sd.get("edition")
                edition_date_text = sd.get("edition_date_text")
                if not edition_value and not edition_date_text:
                    return m.group(0)
                args = split_ifc_args(m.group(3))
                if len(args) < 3:
                    return m.group(0)

                new_edition = ifc_literal(edition_value)
                new_edition_date = ifc_literal(edition_date_text)
                old_edition = args[1].strip()
                old_edition_date = args[2].strip()
                if old_edition == new_edition and old_edition_date == new_edition_date:
                    return m.group(0)

                args[1] = new_edition
                args[2] = new_edition_date
                cls_subs_list[0] += 1
                return m.group(1) + ", ".join(args) + m.group(4)

            cls_subs_list = [0]
            updated = cls_line_regex.sub(repl_cls, updated)
            cls_subs = cls_subs_list[0]

        # 1b) Update IfcClassificationReference.Location using the correct system's URI map.
        ref_updates = [0]
        if has_ref_uri_maps and cls_id_to_system:
            ref_regex = re.compile(
                r"(^\s*#\d+\s*=\s*IFCCLASSIFICATIONREFERENCE\s*\()(.*?)(\)\s*;)",
                re.IGNORECASE | re.MULTILINE
            )

            def repl_ref(m):
                args = split_ifc_args(m.group(2))
                if len(args) < 4:
                    return m.group(0)
                ref_source = args[3].strip()
                source_id  = ref_source.lstrip("#") if ref_source.startswith("#") else ""
                sd = cls_id_to_system.get(source_id)
                if sd is None:
                    return m.group(0)
                by_code = sd.get("class_uri_by_code", {})
                by_pair = sd.get("class_uri_by_pair", {})
                ident    = normalize_text(ifc_unquote(args[1]))
                ref_name = normalize_text(ifc_unquote(args[2]))
                new_loc = None
                if ident and ref_name:
                    new_loc = by_pair.get((ident, ref_name))
                if (not new_loc) and ident:
                    new_loc = by_code.get(ident)
                if not new_loc:
                    return m.group(0)
                args[0] = ifc_literal(new_loc)
                ref_updates[0] += 1
                return m.group(1) + ", ".join(args) + m.group(3)

            updated = ref_regex.sub(repl_ref, updated)

        # 2) Attach every IfcClassification to IfcProject when missing.
        proj_match = re.search(r"#(\d+)\s*=\s*IFCPROJECT\s*\(", updated, re.IGNORECASE)
        rels_added   = 0
        rel_guid_fixed = False
        if proj_match:
            proj_id = proj_match.group(1)
            all_cls_ids = [cm.group(2) for cm in cls_line_regex.finditer(updated)]
            rel_line_regex = re.compile(
                r"(^\s*#\d+\s*=\s*IFCRELASSOCIATESCLASSIFICATION\s*\()(.*?)(\)\s*;)",
                re.IGNORECASE | re.MULTILINE
            )
            existing_proj_cls_ids = set()

            def repl_rel(m):
                args = split_ifc_args(m.group(2))
                if len(args) < 6:
                    return m.group(0)
                if ("#" + proj_id) not in args[4]:
                    return m.group(0)
                cls_ref = args[5].strip()
                if cls_ref.startswith("#"):
                    existing_proj_cls_ids.add(cls_ref.lstrip("#"))
                if not is_ifc_guid_token(args[0]):
                    args[0] = ifc_literal(new_ifc_guid_22())
                    rel_guid_fixed_list[0] = True
                    return m.group(1) + ", ".join(args) + m.group(3)
                return m.group(0)

            rel_guid_fixed_list = [False]
            updated = rel_line_regex.sub(repl_rel, updated)
            rel_guid_fixed = rel_guid_fixed_list[0]

            missing_cls_ids = [cid for cid in all_cls_ids if cid not in existing_proj_cls_ids]
            if missing_cls_ids:
                ids = [int(x) for x in re.findall(r"#(\d+)\s*=", updated)]
                next_id = (max(ids) + 1) if ids else 1
                new_rel_lines = []
                for cls_id_str in missing_cls_ids:
                    guid22 = new_ifc_guid_22()
                    new_rel_lines.append(
                        "#{}=IFCRELASSOCIATESCLASSIFICATION('{}',$,'Classification System',$,(#{}),#{});".format(
                            next_id, guid22, proj_id, cls_id_str
                        )
                    )
                    next_id += 1
                new_rel = "\n".join(new_rel_lines) + "\n"
                endsec_idx = updated.upper().rfind("ENDSEC;")
                if endsec_idx > -1:
                    updated = updated[:endsec_idx] + new_rel + updated[endsec_idx:]
                else:
                    updated += "\n" + new_rel
                rels_added += len(missing_cls_ids)

        if updated != content:
            with open(ifc_path, "w") as f:
                f.write(updated)

        info = []
        if cls_subs:
            info.append("IfcClassification edition/date updated ({} entries)".format(cls_subs))
        if ref_updates[0] > 0:
            info.append("IfcClassificationReference.Location set ({})".format(ref_updates[0]))
        if rel_guid_fixed:
            info.append("IfcRelAssociatesClassification GlobalId corrected")
        if rels_added:
            info.append("{} IfcClassification(s) attached to IfcProject".format(rels_added))
        if not info:
            info.append("No IFC post-processing changes required")
        return (True, "; ".join(info))

    def patch_element_params(all_systems_data, element_class_map):
        # SUPPORTED_PARAMS maps position → shared-parameter name used by the IFC exporter.
        SUPPORTED_PARAMS = ["ClassificationCode", "ClassificationCode(2)", "ClassificationCode(3)", "ClassificationCode(4)", "ClassificationCode(5)"]
        # (dict_name, dict_uri) → the shared parameter it should be written to
        dict_to_param = {}
        for sd in all_systems_data:
            pname = sd.get("param_name")
            if not pname:
                continue
            d_name_key = normalize_text(sd.get("dict_name", ""))
            d_uri_key = normalize_text(sd.get("dict_uri", ""))
            if not d_name_key:
                continue
            dict_to_param[(d_name_key, d_uri_key)] = pname
            if d_uri_key:
                # Fallback key for entries without URI value.
                dict_to_param[(d_name_key, "")] = pname

        bound_cache = set()  # (category_id_int, parameter_name)

        def set_all_named_params(elem, pname, value):
            try:
                for param in elem.Parameters:
                    try:
                        if param.Definition and param.Definition.Name == pname and not param.IsReadOnly:
                            param.Set(value)
                    except Exception:
                        pass
            except Exception:
                pass

        for eid_int, cls_items in element_class_map.items():
            try:
                elem = doc.GetElement(DB.ElementId(int(eid_int)))
                if not elem or not elem.IsValidObject:
                    continue
                if not cls_items:
                    continue
                cat = elem.Category
                if not cat:
                    continue

                # Ensure all supported params are bound for this category.
                cat_id_int = cat.Id.IntegerValue
                for pname in SUPPORTED_PARAMS:
                    cache_key = (cat_id_int, pname)
                    if cache_key in bound_cache:
                        continue
                    ensure_bound(cat, pname, ifc_builtin_sp)
                    bound_cache.add(cache_key)

                # Reset the export slots to avoid stale values from previous assignments.
                for pname in SUPPORTED_PARAMS:
                    set_all_named_params(elem, pname, "")

                # Write every classification item of this element to the mapped slot.
                for cls_item in cls_items:
                    dict_name_val = cls_item.get("dict_name", "")
                    dict_uri_val = cls_item.get("dict_uri", "")
                    code_val = cls_item.get("code", "")
                    name_val = cls_item.get("name", "")
                    if not code_val:
                        continue
                    d_name_key = normalize_text(dict_name_val)
                    d_uri_key = normalize_text(dict_uri_val)
                    target_param = dict_to_param.get((d_name_key, d_uri_key))
                    if not target_param:
                        target_param = dict_to_param.get((d_name_key, ""))
                    if not target_param:
                        continue
                    val = u"{} : {}".format(code_val, name_val) if name_val else code_val
                    set_all_named_params(elem, target_param, val)
            except Exception:
                pass

    # 1) Collect dictionary systems stored in the model.
    legacy_schema = Schema.Lookup(System.Guid(cls_schema_guid))
    multi_schema = Schema.Lookup(System.Guid(cls_multi_schema_guid))
    if legacy_schema is None and multi_schema is None:
        forms.alert(
            "No external classification schema found in the model.\n"
            "Please classify elements first using 'Add Classification'.",
            title="Export IFC"
        )
        return

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

    systems = {}  # (dict_name, dict_uri) -> count
    classified_entries = []  # (dict_name, dict_uri, code, name, class_uri)
    element_class_map = {}  # element id int -> list of classification items

    all_ids = list(DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElementIds())
    for eid in all_ids:
        try:
            elem = doc.GetElement(eid)
            if not elem or not elem.IsValidObject:
                continue
            elem_items = read_element_classifications(
                elem,
                legacy_schema=legacy_schema,
                legacy_fields=legacy_fields,
                multi_schema=multi_schema,
                multi_items_field=multi_items_field
            )
            if not elem_items:
                continue

            element_class_map[eid.IntegerValue] = list(elem_items)

            for item in elem_items:
                dict_name = item.get("dict_name", "")
                dict_uri = item.get("dict_uri", "")
                code_val = item.get("code", "")
                name_val = item.get("name", "")
                class_uri_val = item.get("class_uri", "")
                if not dict_name:
                    continue
                classified_entries.append((dict_name, dict_uri, code_val, name_val, class_uri_val))
                key = (dict_name, dict_uri)
                systems[key] = systems.get(key, 0) + 1
        except Exception:
            pass

    if not systems:
        forms.alert(
            "No classified elements found.\n"
            "Please classify elements first using 'Add Classification'.",
            title="Export IFC"
        )
        return

    # 2) Build classification data for ALL systems. Each system gets its own IFCClassification
    #    DataStorage entry and a dedicated shared parameter slot (ClassificationCode /
    #    ClassificationCode(2) through ClassificationCode(5)). Up to 5 systems are supported.
    PARAM_NAMES = ["ClassificationCode", "ClassificationCode(2)", "ClassificationCode(3)", "ClassificationCode(4)", "ClassificationCode(5)"]
    sorted_systems = sorted(systems.items(), key=lambda kv: kv[1], reverse=True)
    if len(sorted_systems) > len(PARAM_NAMES):
        forms.alert(
            u"More than {} classification systems found in the model ({} in total).\n"
            u"Only the {} most-used systems will be included in the IFC export.".format(
                len(PARAM_NAMES), len(sorted_systems),
                len(PARAM_NAMES), len(PARAM_NAMES)),
            title="Export IFC",
            warn_icon=True
        )
        sorted_systems = sorted_systems[:len(PARAM_NAMES)]

    # Build URI lookup maps in a single pass over classified elements.
    uri_maps_by_system = {}  # (dict_name, dict_uri) -> {"by_code": {}, "by_pair": {}}
    for dict_name, dict_uri, code_val, name_val, class_uri_val in classified_entries:
        code_val = (code_val or "").strip()
        class_uri_val = (class_uri_val or "").strip()
        if not code_val or not class_uri_val:
            continue
        key = (dict_name, dict_uri)
        uri_maps = uri_maps_by_system.get(key)
        if uri_maps is None:
            uri_maps = {"by_code": {}, "by_pair": {}}
            uri_maps_by_system[key] = uri_maps
        code_key = normalize_text(code_val)
        name_key = normalize_text(name_val)
        uri_maps["by_code"][code_key] = class_uri_val
        if name_key:
            uri_maps["by_pair"][(code_key, name_key)] = class_uri_val

    all_systems_data = []
    for i, ((d_name, d_uri), _) in enumerate(sorted_systems):
        edition, edition_date_parts, edition_date_text = get_dictionary_metadata(d_name, d_uri)
        uri_maps = uri_maps_by_system.get((d_name, d_uri), {})
        uri_by_code = uri_maps.get("by_code", {})
        uri_by_pair = uri_maps.get("by_pair", {})
        
        if not uri_by_code:
            forms.alert(
                u"No classified elements found for system '{}'.\n"
                u"This system will be exported without element references.\n\n"
                u"URI: {}".format(d_name or "(empty)", d_uri or "(empty)"),
                title="Export IFC",
                warn_icon=True
            )
        all_systems_data.append({
            "dict_name":          d_name,
            "dict_uri":           d_uri,
            "edition":            edition,
            "edition_date_parts": edition_date_parts,
            "edition_date_text":  edition_date_text,
            "class_uri_by_code":  dict(uri_by_code),
            "class_uri_by_pair":  dict(uri_by_pair),
            "param_name":         PARAM_NAMES[i],
        })

    # 3) Choose export target in a single save dialog.
    out_path = choose_ifc_output_path(doc.Title)
    if not out_path:
        return
    export_folder = os.path.dirname(out_path)
    base_name = os.path.splitext(os.path.basename(out_path))[0]

    # 4) Update IFC-relevant data in the model.
    # Exporter reads IfcClassification from official IFCClassification extensible storage.
    try:
        tx_patch = DB.Transaction(doc, "Prepare External Classification IFC Parameters")
        tx_patch.Start()
        patch_element_params(all_systems_data, element_class_map)
        tx_patch.Commit()

        tx_storage = DB.Transaction(doc, "Prepare External Classification IFC Storage")
        tx_storage.Start()
        write_ifc_classification_storage(all_systems_data)
        tx_storage.Commit()
    except Exception as prep_ex:
        try:
            if 'tx_storage' in locals() and tx_storage.GetStatus() == DB.TransactionStatus.Started:
                tx_storage.RollBack()
        except Exception:
            pass
        try:
            if 'tx_patch' in locals() and tx_patch.GetStatus() == DB.TransactionStatus.Started:
                tx_patch.RollBack()
        except Exception:
            pass
        forms.alert(
            u"IFC export preparation failed:\n{}".format(prep_ex),
            title="Export IFC"
        )
        return

    # 5) Set IFC export options.
    # Let the user choose one of the IFC Exporter add-in profiles. All settings
    # (IFC version, view filter, property sets, etc.) come from that profile.
    opts = _choose_ifc_export_options()
    if opts is None:
        return  # user cancelled profile selection

    # 6) Execute export.
    # Some IFC exporter versions require a temporary open transaction.
    try:
        ok = doc.Export(export_folder, base_name, opts)
    except Exception as ex:
        message = str(ex)
        if "no open transaction" in message.lower() or "modifying is forbidden" in message.lower():
            try:
                tx = DB.Transaction(doc, "External Classification IFC Export")
                tx.Start()
                ok = doc.Export(export_folder, base_name, opts)
                tx.RollBack()
            except Exception as tx_ex:
                forms.alert(
                    u"IFC export failed:\n{}\n\nFallback with transaction also failed:\n{}".format(ex, tx_ex),
                    title="Export IFC"
                )
                return
        else:
            forms.alert(
                u"IFC export failed:\n{}".format(ex),
                title="Export IFC"
            )
            return

    if ok:
        pp_ok, pp_msg = postprocess_ifc_output(
            out_path,
            all_systems_data
        )
        if not pp_ok:
            forms.alert(
                u"IFC export failed during post-processing:\n{}".format(pp_msg),
                title="Export IFC"
            )
            return
        forms.alert("IFC export completed successfully.", title="Export IFC")
    else:
        forms.alert(
            "IFC export failed without a specific error message.",
            title="Export IFC"
        )


main()
