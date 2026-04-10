# -*- coding: utf-8 -*-
__title__ = "Export\nIFC"
__doc__ = "Exports IFC with external classification."

import os
from pyrevit import forms


def main():
    import re
    import sys
    import os as _os_check
    
    # Early compatibility check
    try:
        script_dir = _os_check.path.dirname(__file__)
        parent_dir = _os_check.path.dirname(script_dir)
        if parent_dir not in sys.path:
            sys.path.insert(0, parent_dir)
        import revit_compat as _compat_early
        if not _compat_early.ensure_revit_version(min_version=2022):
            return  # User cancelled or incompatible version
    except Exception as early_ex:
        forms.alert("Compatibility check failed: {}".format(str(early_ex)), title="Error")
        return
    import uuid
    import datetime
    import json
    import sys
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
    script_dir = os.path.dirname(__file__)
    parent_dir = os.path.dirname(script_dir)
    add_cls_dir = os.path.join(parent_dir, "02_AddClassification.pushbutton")
    if add_cls_dir not in sys.path:
        sys.path.insert(0, add_cls_dir)
    
    # Add panel root to sys.path for revit_compat module
    panel_root = parent_dir
    if panel_root not in sys.path:
        sys.path.insert(0, panel_root)
    
    import shared_text_utils as text_utils
    import revit_compat as compat

    try:
        text_type = unicode
    except NameError:
        text_type = str

    safe_unicode = text_utils.safe_unicode
    safe_query_value = text_utils.safe_query_value
    _ifc_literal_ascii = text_utils.ifc_literal_ascii

    doc = __revit__.ActiveUIDocument.Document

    url_schema_guid = "12345678-1234-1234-1234-1234567890ab"
    url_field_name = "APIUrl"
    cls_schema_guid = "ABCDEF12-3456-7890-ABCD-EF1234567890"
    cls_multi_schema_guid = "ABCDEF12-3456-7890-ABCD-EF1234567891"
    ifc_cls_schema_guid = "9A5A28C2-DDAC-4828-8B8A-3EE97118017A"
    # Dynamischer Pfad für kompatible Revit-Versionen (2022-2025+)
    ifc_builtin_sp = compat.get_ifc_shared_parameters_path() or r"C:\Program Files\Autodesk\Revit 2025\IFC Shared Parameters-RevitIFCBuiltIn_ALL.txt"

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
                safe[safe_query_value(k)] = safe_query_value(v)
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
                                    "dict_name": safe_unicode(p.get("dict_name") or "").strip(),
                                    "dict_uri": safe_unicode(p.get("dict_uri") or "").strip(),
                                    "code": safe_unicode(p.get("code") or "").strip(),
                                    "name": safe_unicode(p.get("name") or "").strip(),
                                    "class_uri": safe_unicode(p.get("class_uri") or "").strip(),
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
                    code_val = ent.Get[System.String](legacy_fields["Code"]) or ""
                    if code_val:
                        items.append({
                            "dict_name": safe_unicode(ent.Get[System.String](legacy_fields["DictionaryName"]) or ""),
                            "dict_uri": safe_unicode(ent.Get[System.String](legacy_fields["DictionaryUri"]) or ""),
                            "code": safe_unicode(code_val),
                            "name": safe_unicode(ent.Get[System.String](legacy_fields["Name"]) or ""),
                            "class_uri": safe_unicode(ent.Get[System.String](legacy_fields["ClassUri"]) or ""),
                            "parent_class_uri": "",
                            "parent_class_code": "",
                            "parent_class_name": "",
                            "ancestor_classes": [],
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

    class ExportSettingsDialog(forms.WPFWindow):
        def __init__(self, config_names, info_url):
            xaml_path = os.path.join(os.path.dirname(__file__), "export_settings_dialog.xaml")
            forms.WPFWindow.__init__(self, xaml_path)
            self._info_url = info_url
            self.selected_profile = None
            self.selected_mode = "flat"

            self.ProfileComboBox.Items.Clear()
            if config_names:
                for name in config_names:
                    self.ProfileComboBox.Items.Add(name)
                self.ProfileComboBox.SelectedIndex = 0
            else:
                self.ProfileComboBox.Items.Add("Default IFC 4")
                self.ProfileComboBox.SelectedIndex = 0
                self.ProfileComboBox.IsEnabled = False

            self.FlatRadio.IsChecked = True

            # Bind in code as a reliable fallback for environments where XAML
            # event hookup can be inconsistent.
            try:
                self.InfoBtn.Click += self.InfoBtn_Click
            except Exception:
                pass

        def InfoBtn_Click(self, sender, args):
            opened = False
            try:
                os.startfile(self._info_url)
                opened = True
            except Exception:
                try:
                    psi = System.Diagnostics.ProcessStartInfo(self._info_url)
                    psi.UseShellExecute = True
                    System.Diagnostics.Process.Start(psi)
                    opened = True
                except Exception:
                    opened = False

            if not opened:
                forms.alert(
                    u"Could not open documentation link:\n{}".format(self._info_url),
                    title="Export IFC",
                    warn_icon=True
                )

        def ContinueBtn_Click(self, sender, args):
            try:
                if self.ProfileComboBox.IsEnabled and self.ProfileComboBox.SelectedItem is not None:
                    self.selected_profile = safe_unicode(self.ProfileComboBox.SelectedItem)
                else:
                    self.selected_profile = None
            except Exception:
                self.selected_profile = None

            self.selected_mode = "hierarchical" if bool(self.HierarchicalRadio.IsChecked) else "flat"
            self.DialogResult = True
            self.Close()

        def CancelBtn_Click(self, sender, args):
            self.DialogResult = False
            self.Close()

    def _show_export_settings_dialog(config_names):
        info_url = "https://standards.buildingsmart.org/IFC/RELEASE/IFC4_3/HTML/lexical/IfcClassificationReference.htm"

        dlg = ExportSettingsDialog(config_names, info_url)
        result = dlg.ShowDialog()
        if not result:
            return (None, None)
        return (dlg.selected_profile, dlg.selected_mode)

    def _choose_ifc_export_options_and_mode():
        """
        Tries to find the IFC Exporter add-in (IFCExporterUIOverride.dll, already
        loaded in Revit's process) and offers all available export configurations
        (built-in + user-saved profiles).
        Applies the chosen profile to a fresh IFCExportOptions instance via
        IFCExportConfiguration.UpdateOptions().

        Falls back to basic IFC 4 options when the add-in is unavailable.
        Returns (None, None) when the user cancels.
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

        config_names = []
        name_to_config = {}

        if config_maps_type is not None:
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
            except Exception:
                config_names = []
                name_to_config = {}

        selected_name, export_mode = _show_export_settings_dialog(config_names)
        if export_mode is None:
            return (None, None)

        if not config_names:
            try:
                base_opts.FileVersion = DB.IFCVersion.IFC4
            except Exception:
                pass
        elif selected_name:
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

        return (base_opts, export_mode)

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

    def inject_properties_into_ifc(ifc_path, element_properties_map):
        """
        Inject stored class properties directly into the IFC as
        IfcPropertySet + IfcPropertySingleValue + IfcRelDefinesByProperties.
        """
        if not element_properties_map:
            return (False, "No properties found on classified elements")

        if not os.path.exists(ifc_path):
            return (False, "IFC file not found")

        with open(ifc_path, "r") as f:
            content = f.read()

        endsec_idx = content.upper().rfind("ENDSEC;")
        if endsec_idx < 0:
            return (False, "Invalid IFC structure: ENDSEC not found")

        ids = [int(x) for x in re.findall(r"#(\d+)\s*=", content)]
        next_id = (max(ids) + 1) if ids else 1

        target_elem_ids = set(int(k) for k in element_properties_map.keys())

        def _to_unicode(value):
            try:
                txt = text_type(value or "")
            except Exception:
                txt = "{}".format(value or "")

            try:
                if isinstance(txt, unicode):
                    return txt
            except Exception:
                pass

            try:
                return txt.decode("utf-8")
            except Exception:
                pass
            try:
                return txt.decode("cp1252")
            except Exception:
                pass
            try:
                return txt.decode("latin-1")
            except Exception:
                pass
            try:
                return unicode(txt)
            except Exception:
                return u""

        def _decode_storage_escapes(value):
            txt = _to_unicode(value)
            try:
                if "\\x" in txt or "\\u" in txt or "\\U" in txt:
                    try:
                        decoded = txt.encode("utf-8").decode("unicode_escape")
                    except Exception:
                        decoded = str(txt).decode("unicode_escape")
                    if decoded:
                        txt = _to_unicode(decoded)
            except Exception:
                pass
            return txt

        # Best-effort mapping: Revit element id is typically at end of IFC element Name.
        # Example: "Basic Wall:Generic - 200mm:123456" -> element id 123456
        entity_by_elem = {}
        line_regex = re.compile(r"^\s*#(\d+)\s*=\s*IFC[A-Z0-9_]+\s*\((.*)\)\s*;\s*$", re.IGNORECASE)
        for line in content.splitlines():
            m = line_regex.match(line)
            if not m:
                continue
            ent_id = int(m.group(1))
            args = split_ifc_args(m.group(2))
            if len(args) < 3:
                continue
            name_txt = ifc_unquote(args[2])
            mm = re.search(r"(?:^|[:#\s])(\d+)\s*$", name_txt or "")
            if not mm:
                continue
            try:
                rid = int(mm.group(1))
            except Exception:
                continue
            if rid in target_elem_ids:
                entity_by_elem.setdefault(rid, []).append(ent_id)

        if not entity_by_elem:
            return (False, "Could not map Revit element ids to IFC entities")

        new_lines = []
        rel_count = 0
        pset_count = 0
        prop_count = 0

        for elem_id, class_map in element_properties_map.items():
            try:
                elem_id_int = int(elem_id)
            except Exception:
                continue

            target_entities = entity_by_elem.get(elem_id_int, [])
            if not target_entities:
                continue

            if not isinstance(class_map, dict):
                continue

            for class_uri, pdata in class_map.items():
                psets = list((pdata or {}).get("property_sets", []) or [])
                for pset in psets:
                    pset_name = (pset.get("name") or "PropertySet").strip()
                    props = list(pset.get("properties", []) or [])
                    single_ids = []

                    for prop in props:
                        if not isinstance(prop, dict):
                            continue
                        enabled = bool(prop.get("enabled", False))
                        raw_val = prop.get("value", "")
                        if (not enabled) and raw_val in (None, ""):
                            continue

                        prop_name = (prop.get("name") or "Property").strip()
                        prop_name_lit = _ifc_literal_ascii(prop_name)
                        value_lit = _ifc_literal_ascii(raw_val)
                        nominal = "$" if value_lit == "$" else "IFCTEXT({})".format(value_lit)

                        sid = next_id
                        next_id += 1
                        new_lines.append(
                            "#{}=IFCPROPERTYSINGLEVALUE({},$,{},{}) ;".format(sid, prop_name_lit, nominal, "$")
                        )
                        single_ids.append(sid)
                        prop_count += 1

                    if not single_ids:
                        continue

                    pset_id = next_id
                    next_id += 1
                    pset_name_lit = _ifc_literal_ascii(pset_name)
                    has_props = "(" + ",".join("#{}".format(i) for i in single_ids) + ")"
                    new_lines.append(
                        "#{}=IFCPROPERTYSET({},$,{},{},{});".format(
                            pset_id,
                            _ifc_literal_ascii(new_ifc_guid_22()),
                            pset_name_lit,
                            "$",
                            has_props
                        )
                    )
                    pset_count += 1

                    for ent_id in target_entities:
                        rel_id = next_id
                        next_id += 1
                        new_lines.append(
                            "#{}=IFCRELDEFINESBYPROPERTIES({},$,$,$,(#{}),#{});".format(
                                rel_id,
                                _ifc_literal_ascii(new_ifc_guid_22()),
                                ent_id,
                                pset_id
                            )
                        )
                        rel_count += 1

        if not new_lines:
            return (False, "No enabled/non-empty properties to inject")

        injected = "\n".join(new_lines) + "\n"
        updated = content[:endsec_idx] + injected + content[endsec_idx:]

        with open(ifc_path, "w") as f:
            f.write(updated)

        return (True, "Injected {} property values in {} property sets ({} relations)".format(prop_count, pset_count, rel_count))

    def inject_classifications_into_ifc(ifc_path, element_class_map, all_systems_data, export_mode="flat"):
        """Inject additional missing class associations directly into the IFC file."""
        if not element_class_map:
            return (False, "No classified elements found")
        if not os.path.exists(ifc_path):
            return (False, "IFC file not found")

        with open(ifc_path, "r") as f:
            content = f.read()

        endsec_idx = content.upper().rfind("ENDSEC;")
        if endsec_idx < 0:
            return (False, "Invalid IFC structure: ENDSEC not found")

        ids = [int(x) for x in re.findall(r"#(\d+)\s*=", content)]
        next_id = (max(ids) + 1) if ids else 1

        def _ifc_guid_22():
            alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$"
            num = uuid.uuid4().int
            out = []
            for _ in range(22):
                out.append(alphabet[num & 63])
                num >>= 6
            return "".join(reversed(out))

        # Map Revit element ids to IFC entity ids by the exported element name suffix.
        entity_by_elem = {}
        line_regex = re.compile(r"^\s*#(\d+)\s*=\s*IFC[A-Z0-9_]+\s*\((.*)\)\s*;\s*$", re.IGNORECASE)
        for line in content.splitlines():
            m = line_regex.match(line)
            if not m:
                continue
            ent_id = int(m.group(1))
            args = split_ifc_args(m.group(2))
            if len(args) < 3:
                continue
            name_txt = ifc_unquote(args[2])
            mm = re.search(r"(?:^|[:#\s])(\d+)\s*$", name_txt or "")
            if not mm:
                continue
            try:
                rid = int(mm.group(1))
            except Exception:
                continue
            entity_by_elem.setdefault(rid, []).append(ent_id)

        if not entity_by_elem:
            return (False, "Could not map Revit element ids to IFC entities")

        # Find IFC classification entity ids for each classification system.
        cls_source_ids = {}
        cls_regex = re.compile(r"^\s*#(\d+)\s*=\s*IFCCLASSIFICATION\s*\((.*)\)\s*;\s*$", re.IGNORECASE)
        for line in content.splitlines():
            m = cls_regex.match(line)
            if not m:
                continue
            cls_id = m.group(1)
            args = split_ifc_args(m.group(2))
            if len(args) < 6:
                continue
            cls_name = ifc_unquote(args[3])
            cls_loc = ifc_unquote(args[5])
            for sd in all_systems_data:
                d_name = sd.get("dict_name", "")
                d_uri = sd.get("dict_uri", "")
                if (d_uri and normalize_text(cls_loc) == normalize_text(d_uri)) or (d_name and normalize_text(cls_name) == normalize_text(d_name)):
                    cls_source_ids[(normalize_text(d_name), normalize_text(d_uri))] = cls_id
                    break

        system_by_key = {}
        for sd in all_systems_data:
            k = (normalize_text(sd.get("dict_name", "")), normalize_text(sd.get("dict_uri", "")))
            system_by_key[k] = sd

        existing_ref_by_key = {}
        ref_line_regex = re.compile(r"^\s*#(\d+)\s*=\s*IFCCLASSIFICATIONREFERENCE\s*\((.*)\)\s*;\s*$", re.IGNORECASE)
        for line in content.splitlines():
            m = ref_line_regex.match(line)
            if not m:
                continue
            ref_id = m.group(1)
            args = split_ifc_args(m.group(2))
            if len(args) < 4:
                continue
            ref_uri = normalize_text(ifc_unquote(args[0]))
            src = args[3].strip()
            src_id = src.lstrip("#") if src.startswith("#") else src
            if ref_uri and src_id:
                existing_ref_by_key[(ref_uri, src_id)] = ref_id

        existing_rel_pairs = set()
        rel_line_regex = re.compile(r"^\s*#(\d+)\s*=\s*IFCRELASSOCIATESCLASSIFICATION\s*\((.*)\)\s*;\s*$", re.IGNORECASE)
        for line in content.splitlines():
            m = rel_line_regex.match(line)
            if not m:
                continue
            args = split_ifc_args(m.group(2))
            if len(args) < 6:
                continue
            related_arg = args[4] or ""
            src = args[5].strip()
            src_id = src.lstrip("#") if src.startswith("#") else src
            if not src_id:
                continue
            for ent in re.findall(r"#(\d+)", related_arg):
                existing_rel_pairs.add((ent, src_id))

        # Inject additional references and relations while preserving existing export output.
        new_lines = []
        rel_count = 0
        ref_count = 0
        for elem_id, cls_items in element_class_map.items():
            try:
                elem_id_int = int(elem_id)
            except Exception:
                continue
            target_entities = entity_by_elem.get(elem_id_int, [])
            if not target_entities:
                continue
            if not isinstance(cls_items, list):
                continue

            for item in cls_items:
                sys_key = (normalize_text(item.get("dict_name", "")), normalize_text(item.get("dict_uri", "")))
                src_id = cls_source_ids.get(sys_key)
                if not src_id:
                    sd = system_by_key.get(sys_key)
                    if sd:
                        src_id = str(next_id)
                        next_id += 1
                        cls_source_ids[sys_key] = src_id
                        cls_edition = ifc_literal(sd.get("edition"))
                        cls_edition_date = ifc_literal(sd.get("edition_date_text"))
                        new_lines.append(
                            "#{}=IFCCLASSIFICATION($,{},{},{},$,{},$);".format(
                                src_id,
                                cls_edition,
                                cls_edition_date,
                                _ifc_literal_ascii(sd.get("dict_name", "")),
                                _ifc_literal_ascii(sd.get("dict_uri", ""))
                            )
                        )
                    else:
                        continue

                class_uri_val = safe_unicode(item.get("class_uri", "")).strip()
                if not class_uri_val:
                    continue

                chain = []
                if export_mode == "hierarchical":
                    raw_ancestors = item.get("ancestor_classes") or []
                    if isinstance(raw_ancestors, list):
                        for ancestor in raw_ancestors:
                            if not isinstance(ancestor, dict):
                                continue
                            a_uri = safe_unicode(ancestor.get("uri", "")).strip()
                            if not a_uri:
                                continue
                            chain.append({
                                "uri": a_uri,
                                "code": safe_unicode(ancestor.get("code", "")).strip(),
                                "name": safe_unicode(ancestor.get("name", "")).strip(),
                            })

                    if not chain:
                        p_uri = safe_unicode(item.get("parent_class_uri", "")).strip()
                        if p_uri:
                            chain.append({
                                "uri": p_uri,
                                "code": safe_unicode(item.get("parent_class_code", "")).strip(),
                                "name": safe_unicode(item.get("parent_class_name", "")).strip(),
                            })

                # Leaf is always present. In flat mode this is the only node.
                chain.append({
                    "uri": class_uri_val,
                    "code": safe_unicode(item.get("code", "")).strip(),
                    "name": safe_unicode(item.get("name", "")).strip(),
                })

                chain_src = str(src_id)
                leaf_ref_id = None
                for node in chain:
                    node_uri = safe_unicode(node.get("uri", "")).strip()
                    if not node_uri:
                        continue

                    ref_key = (normalize_text(node_uri), chain_src)
                    ref_id = existing_ref_by_key.get(ref_key)

                    if not ref_id:
                        ref_id = str(next_id)
                        next_id += 1
                        node_name = safe_unicode(node.get("name", "")).strip()
                        node_code = safe_unicode(node.get("code", "")).strip()
                        new_lines.append(
                            "#{}=IFCCLASSIFICATIONREFERENCE({},{},{},#{},$,$);".format(
                                ref_id,
                                _ifc_literal_ascii(node_uri),
                                _ifc_literal_ascii(node_code),
                                _ifc_literal_ascii(node_name),
                                chain_src
                            )
                        )
                        existing_ref_by_key[ref_key] = ref_id
                        ref_count += 1

                    leaf_ref_id = ref_id
                    chain_src = ref_id

                if not leaf_ref_id:
                    continue

                for ent_id in target_entities:
                    rel_key = (str(ent_id), str(leaf_ref_id))
                    if rel_key in existing_rel_pairs:
                        continue
                    rel_id = next_id
                    next_id += 1
                    new_lines.append(
                        "#{}=IFCRELASSOCIATESCLASSIFICATION({},$,$,$,(#{}),#{});".format(
                            rel_id,
                            _ifc_literal_ascii(_ifc_guid_22()),
                            ent_id,
                            leaf_ref_id
                        )
                    )
                    existing_rel_pairs.add(rel_key)
                    rel_count += 1

        if not new_lines:
            return (False, "No additional classifications to inject")

        injected = "\n".join(new_lines) + "\n"
        updated = content[:endsec_idx] + injected + content[endsec_idx:]

        with open(ifc_path, "w") as f:
            f.write(updated)

        return (True, "Injected {} class references in {} relations".format(ref_count, rel_count))

    def patch_element_params(all_systems_data, element_class_map, write_values=True):
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

                # In hierarchical mode we only clear slots to avoid duplicate flat links.
                if not write_values:
                    continue

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
    element_properties_map = {}  # element id int -> dict of class_uri -> property_sets

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
            
            # Try to load properties if available
            try:
                import sys
                script_dir = os.path.dirname(__file__)
                parent_dir = os.path.dirname(script_dir)
                add_cls_dir = os.path.join(parent_dir, "02_AddClassification.pushbutton")
                if add_cls_dir not in sys.path:
                    sys.path.insert(0, add_cls_dir)
                
                import properties_editor as pe_export
                elem_props = pe_export.load_all_class_properties(elem)
                if elem_props:
                    element_properties_map[eid.IntegerValue] = elem_props
            except Exception:
                # Properties module not available, skip
                pass

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

    # 3) Choose export profile + classification mode in one compact dialog.
    opts, export_mode = _choose_ifc_export_options_and_mode()
    if opts is None or export_mode is None:
        return  # user cancelled

    # 4) Choose export target in a single save dialog.
    out_path = choose_ifc_output_path(doc.Title)
    if not out_path:
        return
    export_folder = os.path.dirname(out_path)
    base_name = os.path.splitext(os.path.basename(out_path))[0]

    # 5) Update IFC-relevant data in the model.
    # Exporter reads IfcClassification from official IFCClassification extensible storage.
    try:
        # Always clear slot parameters to remove stale exporter mappings.
        # Only flat mode writes values back to those slots.
        tx_patch = DB.Transaction(doc, "Prepare External Classification IFC Parameters")
        tx_patch.Start()
        patch_element_params(
            all_systems_data,
            element_class_map,
            write_values=(export_mode == "flat")
        )
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
        
        cls_ok, cls_msg = inject_classifications_into_ifc(
            out_path,
            element_class_map,
            all_systems_data,
            export_mode=export_mode
        )
        props_ok, props_msg = inject_properties_into_ifc(out_path, element_properties_map)

        msg = "IFC export completed successfully."
        forms.alert(msg, title="Export IFC")
    else:
        forms.alert(
            "IFC export failed without a specific error message.",
            title="Export IFC"
        )


main()
