__title__ = "Set URL"
__doc__ = "Set URL for External Classification API"

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, Entity, AccessLevel
from pyrevit import forms
import clr
clr.AddReference('System')
import System

doc = __revit__.ActiveUIDocument.Document

SCHEMA_GUID = System.Guid("12345678-1234-1234-1234-1234567890ab")
SCHEMA_NAME = "DataDictionaryAPI"
FIELD_NAME = "APIUrl"

def get_or_create_schema():
    schema = Schema.Lookup(SCHEMA_GUID)
    if not schema:
        schema_builder = SchemaBuilder(SCHEMA_GUID)
        schema_builder.SetSchemaName(SCHEMA_NAME)
        schema_builder.SetReadAccessLevel(AccessLevel.Public)
        schema_builder.SetWriteAccessLevel(AccessLevel.Public)
        # System.String is the correct .NET type for AddSimpleField
        schema_builder.AddSimpleField(FIELD_NAME, System.String)
        schema = schema_builder.Finish()
    return schema

def set_url():
    url = forms.ask_for_string(
        get_stored_url(),
        "Enter the URL for the External Classification API"
    )
    if url:
        schema = get_or_create_schema()
        field = schema.GetField(FIELD_NAME)
        entity = Entity(schema)
        entity.Set[System.String](field, url)

        with Transaction(doc, "Set External Classification API URL") as t:
            t.Start()
            # URL is stored in ProjectInformation (document level)
            doc.ProjectInformation.SetEntity(entity)
            t.Commit()

        forms.alert("URL set successfully!", title="Success")
    else:
        forms.alert("No URL entered. Operation cancelled.", title="Cancelled")

if __name__ == "__main__":
    def get_stored_url():
        schema = Schema.Lookup(SCHEMA_GUID)
        if schema:
            entity = doc.ProjectInformation.GetEntity(schema)
            if entity.IsValid():
                field = schema.GetField(FIELD_NAME)
                stored = entity.Get[System.String](field)
                if stored:
                    return stored
        return "https://api.bsdd.buildingsmart.org"

    # Ensure default URL exists, but do not show extra popup dialogs.
    schema = get_or_create_schema()
    try:
        existing_entity = doc.ProjectInformation.GetEntity(schema)
        field = schema.GetField(FIELD_NAME)
        stored = ""
        if existing_entity and existing_entity.IsValid():
            stored = existing_entity.Get[System.String](field) or ""
        if not stored:
            entity = Entity(schema)
            entity.Set[System.String](field, "https://api.bsdd.buildingsmart.org")
            with Transaction(doc, "Set Default External Classification API URL") as t:
                t.Start()
                doc.ProjectInformation.SetEntity(entity)
                t.Commit()
    except Exception:
        pass

    set_url()