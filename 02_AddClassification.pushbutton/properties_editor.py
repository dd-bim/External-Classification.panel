# -*- coding: utf-8 -*-
"""
Property management for external classifications.
Handles storage, editing, and export of class properties and property sets.
"""

import json as _json
import shared_text_utils as text_utils


_PROP_SCHEMA_GUID = "FEDCBA98-7654-3210-FEDC-BA9876543211"
_PROP_SCHEMA_NAME = "ExternalClassificationProperties"
_LAST_STORE_ERROR = ""


def get_last_store_error():
    """Return last backend write error message for diagnostics."""
    return _LAST_STORE_ERROR


_safe_unicode = text_utils.safe_unicode


def _normalize_for_json(obj):
    """Recursively normalize values to JSON-safe Python primitives."""
    if obj is None:
        return None
    if isinstance(obj, (bool, int, long, float)):
        return obj
    if isinstance(obj, (unicode, str)):
        return _safe_unicode(obj)
    if isinstance(obj, list):
        return [_normalize_for_json(x) for x in obj]
    if isinstance(obj, tuple):
        return [_normalize_for_json(x) for x in obj]
    if isinstance(obj, dict):
        normalized = {}
        for k, v in obj.items():
            normalized[_safe_unicode(k)] = _normalize_for_json(v)
        return normalized
    return _safe_unicode(obj)


def _safe_json_dumps(obj):
    """Dump JSON using ASCII-safe output to avoid encoding errors."""
    def _ascii_text(value):
        txt = _safe_unicode(value)
        try:
            # Convert to ASCII-safe escaped bytes (e.g. \u00f6 for "ö").
            return txt.encode("unicode_escape")
        except Exception:
            try:
                return txt.encode("ascii", "backslashreplace")
            except Exception:
                return ""

    def _normalize_ascii(obj_inner):
        if obj_inner is None:
            return None
        if isinstance(obj_inner, (bool, int, long, float)):
            return obj_inner
        if isinstance(obj_inner, (unicode, str)):
            return _ascii_text(obj_inner)
        if isinstance(obj_inner, list):
            return [_normalize_ascii(x) for x in obj_inner]
        if isinstance(obj_inner, tuple):
            return [_normalize_ascii(x) for x in obj_inner]
        if isinstance(obj_inner, dict):
            out = {}
            for k, v in obj_inner.items():
                out[_ascii_text(k)] = _normalize_ascii(v)
            return out
        return _ascii_text(obj_inner)

    try:
        normalized = _normalize_for_json(obj)
        return _json.dumps(normalized, ensure_ascii=True)
    except Exception:
        normalized_ascii = _normalize_ascii(obj)
        return _json.dumps(normalized_ascii, ensure_ascii=True)


def _minimal_property_payload(property_data):
    """Reduce payload to essential fields only to keep storage robust."""
    result = {"property_sets": []}
    try:
        psets = list((property_data or {}).get("property_sets", []) or [])
        for pset in psets:
            pset_name = _safe_unicode((pset or {}).get("name", "Other"))
            props_out = []
            for prop in list((pset or {}).get("properties", []) or []):
                try:
                    props_out.append({
                        "uri": _safe_unicode((prop or {}).get("uri", "")),
                        "name": _safe_unicode((prop or {}).get("name", "")),
                        "dataType": _normalize_data_type((prop or {}).get("dataType", None)),
                        "value": _safe_unicode((prop or {}).get("value", "")),
                        "enabled": bool((prop or {}).get("enabled", False)),
                    })
                except Exception:
                    pass
            result["property_sets"].append({
                "name": pset_name,
                "properties": props_out,
            })
    except Exception:
        pass
    return result


def _normalize_data_type(raw_data_type):
    """Normalize bSDD dataType payloads to a readable string."""
    try:
        if raw_data_type is None:
            return "undefined"

        if isinstance(raw_data_type, (unicode, str)):
            val = raw_data_type.strip()
            return val if val else "undefined"

        if isinstance(raw_data_type, dict):
            # bSDD may deliver dataType as object with code/name/value.
            val = raw_data_type.get("code") or raw_data_type.get("name") or raw_data_type.get("value") or ""
            val = _safe_unicode(val).strip()
            return val if val else "undefined"

        if isinstance(raw_data_type, list):
            if not raw_data_type:
                return "undefined"
            first = raw_data_type[0]
            return _normalize_data_type(first)

        val = _safe_unicode(raw_data_type).strip()
        return val if val else "undefined"
    except Exception:
        return "undefined"


def create_property_schema():
    """Create or retrieve the ExtensibleStorage schema for properties."""
    import clr
    clr.AddReference('System')
    import System
    
    from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, AccessLevel
    
    guid = System.Guid(_PROP_SCHEMA_GUID)
    schema = Schema.Lookup(guid)
    if schema:
        return schema
    
    sb = SchemaBuilder(guid)
    sb.SetSchemaName(_PROP_SCHEMA_NAME)
    sb.SetReadAccessLevel(AccessLevel.Public)
    sb.SetWriteAccessLevel(AccessLevel.Public)
    # PropertiesJson stores: { class_uri -> { property_sets } }
    sb.AddSimpleField("PropertiesJson", System.String)
    return sb.Finish()


def _get_properties_field(schema):
    """Resolve the storage field, with fallback for legacy schema variants."""
    try:
        field = schema.GetField("PropertiesJson")
        if field:
            return field
    except Exception:
        pass

    try:
        for field in schema.ListFields():
            if field:
                return field
    except Exception:
        pass

    return None


def store_class_properties(elem, class_uri, property_data):
    """
    Store property set data for a classification on an element.
    
    Args:
        elem: Revit Element
        class_uri: Classification URI (key)
        property_data: dict with structure:
            {
                "property_sets": [
                    {
                        "name": "PropertySet name",
                        "properties": [
                            {
                                "uri": "property_uri",
                                "name": "property_name",
                                "dataType": "string|number|boolean",
                                "value": "user_entered_value",
                                "enabled": True/False,
                                "allowedValues": [...],  # optional
                                "isRequired": True/False,
                            },
                            ...
                        ]
                    }
                ]
            }
    """
    global _LAST_STORE_ERROR
    _LAST_STORE_ERROR = ""

    stage = "init"
    try:
        import clr
        clr.AddReference('System')
        import System
        from Autodesk.Revit.DB.ExtensibleStorage import Entity
        
        stage = "schema"
        schema = create_property_schema()
        if not schema:
            _LAST_STORE_ERROR = "Schema unavailable."
            return False
        
        # Read existing properties; if corrupted, start clean for this element.
        stage = "load-existing"
        all_props = load_all_class_properties(elem)
        if not isinstance(all_props, dict):
            all_props = {}
        
        # Update or add this class's properties
        stage = "normalize"
        class_key = _safe_unicode(class_uri)
        clean_payload = _minimal_property_payload(property_data)
        all_props[class_key] = _normalize_for_json(clean_payload)
        
        # Store back
        stage = "create-entity"
        entity = Entity(schema)
        prop_field = _get_properties_field(schema)
        if prop_field:
            try:
                stage = "json-dumps"
                json_text = _safe_json_dumps(all_props)
            except Exception:
                # Fallback: persist only current class payload if legacy data is broken.
                stage = "json-dumps-fallback"
                fallback_props = {class_key: _normalize_for_json(clean_payload)}
                json_text = _safe_json_dumps(fallback_props)

            try:
                # json_text is ASCII-safe from ensure_ascii=True; avoid extra conversions.
                stage = "entity-set"
                entity.Set[System.String](prop_field, json_text)
                stage = "entity-commit"
                elem.SetEntity(entity)
                return True
            except Exception:
                try:
                    stage = "entity-set-fallback"
                    fallback_text = _json.dumps({class_key: _normalize_for_json(clean_payload)}, ensure_ascii=True)
                    entity.Set[System.String](prop_field, fallback_text)
                    stage = "entity-commit-fallback"
                    elem.SetEntity(entity)
                    return True
                except Exception as write_ex:
                    # Last resort: clear potentially broken entity and write fresh payload.
                    try:
                        stage = "entity-delete-broken"
                        elem.DeleteEntity(schema)
                    except Exception:
                        pass
                    try:
                        stage = "entity-reset-write"
                        clean_entity = Entity(schema)
                        clean_text = _json.dumps({class_key: _normalize_for_json(clean_payload)}, ensure_ascii=True)
                        clean_entity.Set[System.String](prop_field, clean_text)
                        stage = "entity-reset-commit"
                        elem.SetEntity(clean_entity)
                        return True
                    except Exception as reset_ex:
                        _LAST_STORE_ERROR = u"Stage {} | Write failed: {} | Reset failed: {}".format(
                            stage,
                            _safe_unicode(write_ex), _safe_unicode(reset_ex)
                        )[:220]
                        return False
        _LAST_STORE_ERROR = "No writable field found in property schema."
        return False
    except Exception as ex:
        try:
            _LAST_STORE_ERROR = u"Stage {} | {}".format(stage, _safe_unicode(repr(ex)))[:220]
        except Exception:
            _LAST_STORE_ERROR = u"Storage write failed (unreadable exception)."
        return False


def load_class_properties(elem, class_uri):
    """Load property set data for a specific classification."""
    try:
        all_props = load_all_class_properties(elem)
        if class_uri in all_props:
            return all_props.get(class_uri, {})

        want = _safe_unicode(class_uri).strip().lower()
        if not want:
            return {}

        for key, value in all_props.items():
            if _safe_unicode(key).strip().lower() == want:
                return value or {}

        return {}
    except Exception:
        return {}


def load_all_class_properties(elem):
    """Load all property data from element."""
    import clr
    clr.AddReference('System')
    import System
    from Autodesk.Revit.DB.ExtensibleStorage import Schema
    
    try:
        schema = Schema.Lookup(System.Guid(_PROP_SCHEMA_GUID))
        if not schema:
            return {}
        
        entity = elem.GetEntity(schema)
        if not entity or not entity.IsValid():
            return {}
        
        prop_field = _get_properties_field(schema)
        if not prop_field:
            return {}
        
        json_str = entity.Get[System.String](prop_field)
        if not json_str:
            return {}
        
        # Decode safely using normalized unicode only.
        try:
            json_unicode = _safe_unicode(json_str)
            result = _json.loads(json_unicode)
            return result if isinstance(result, dict) else {}
        except Exception:
            return {}
    except Exception:
        return {}


def delete_class_properties(elem, class_uri):
    """Remove property data for a classification when it's deleted."""
    try:
        all_props = load_all_class_properties(elem)
        if not all_props:
            return

        target_key = None
        if class_uri in all_props:
            target_key = class_uri
        else:
            want = _safe_unicode(class_uri).strip().lower()
            for key in all_props.keys():
                if _safe_unicode(key).strip().lower() == want:
                    target_key = key
                    break

        if not target_key:
            return
        
        del all_props[target_key]
        
        if all_props:
            # Store remaining properties
            import clr
            clr.AddReference('System')
            import System
            from Autodesk.Revit.DB.ExtensibleStorage import Entity
            
            schema = create_property_schema()
            if schema:
                entity = Entity(schema)
                prop_field = _get_properties_field(schema)
                if prop_field:
                    entity.Set[System.String](prop_field, _safe_unicode(_safe_json_dumps(all_props)))
                    elem.SetEntity(entity)
        else:
            # No more properties, clear the entity
            try:
                import clr
                clr.AddReference('System')
                import System
                from Autodesk.Revit.DB.ExtensibleStorage import Schema

                schema = Schema.Lookup(System.Guid(_PROP_SCHEMA_GUID))
                if schema:
                    entity_prop = elem.GetEntity(schema)
                    if entity_prop and entity_prop.IsValid():
                        elem.DeleteEntity(schema)
            except Exception:
                pass
    except Exception:
        pass


def build_property_tree(class_data):
    """
    Extract and organize properties from class data into property sets.
    
    Returns:
        list of property sets, each with:
        {
            "name": "PropertySet name",
            "description": "...",
            "properties": [
                {
                    "uri": "property_uri",
                    "name": "property_name",
                    "dataType": "string|number|boolean",
                    "description": "...",
                    "allowedValues": [...],  # or None
                    "isRequired": True/False,
                    "isWritable": True/False,
                    "units": [...],  # or None
                    "example": "...",
                    "value": "",  # user input
                    "enabled": False  # user toggle
                }
            ]
        }
    """
    
    class_props = class_data.get("classProperties", [])
    
    # Group properties by propertySet
    sets_dict = {}
    for prop in class_props:
        set_name = prop.get("propertySet", "Other Properties")
        
        if set_name not in sets_dict:
            sets_dict[set_name] = {
                "name": set_name,
                "description": "",
                "properties": []
            }
        
        prop_item = {
            "uri": prop.get("uri", ""),
            "name": prop.get("name", ""),
            "dataType": _normalize_data_type(prop.get("dataType", None)),
            "description": prop.get("description", ""),
            "definition": prop.get("definition", ""),
            "allowedValues": prop.get("allowedValues", []) or None,
            "isRequired": prop.get("isRequired", False),
            "isWritable": prop.get("isWritable", True),
            "isDynamic": prop.get("isDynamic", False),
            "units": prop.get("units", []) or None,
            "physicalQuantity": prop.get("physicalQuantity", ""),
            "example": prop.get("example", ""),
            "pattern": prop.get("pattern", ""),
            "minInclusive": prop.get("minInclusive", None),
            "maxInclusive": prop.get("maxInclusive", None),
            "value": "",  # will be filled by user
            "enabled": prop.get("isRequired", False)  # required props start as enabled
        }
        
        sets_dict[set_name]["properties"].append(prop_item)
    
    # Return as list, sorted by name
    return sorted(sets_dict.values(), key=lambda x: x["name"])


def serialize_properties_for_storage(property_sets):
    """Convert property tree to JSON-serializable format for storage."""
    return {
        "property_sets": property_sets
    }


def update_property_value(elem, class_uri, pset_name, property_name, new_value):
    """
    Update a single property value in storage.
    
    Args:
        elem: Revit Element
        class_uri: Classification URI
        pset_name: PropertySet name
        property_name: Name of the property to update
        new_value: New value to set
    """
    try:
        all_props = load_all_class_properties(elem)
        if not isinstance(all_props, dict) or not all_props:
            return False

        target_key = None
        if class_uri in all_props:
            target_key = class_uri
        else:
            want = _safe_unicode(class_uri).strip().lower()
            for k in all_props.keys():
                if _safe_unicode(k).strip().lower() == want:
                    target_key = k
                    break

        if not target_key:
            return False

        pdata = all_props.get(target_key, {}) or {}
        psets = list(pdata.get("property_sets", []) or [])
        updated = False

        for pset in psets:
            if (pset.get("name") or "") != (pset_name or ""):
                continue
            for prop in list(pset.get("properties", []) or []):
                if (prop.get("name") or "") == (property_name or ""):
                    prop["value"] = _safe_unicode(new_value)
                    updated = True
                    break
            if updated:
                break

        if not updated:
            return False

        return bool(store_class_properties(elem, target_key, {"property_sets": psets}))

    except Exception:
        return False

