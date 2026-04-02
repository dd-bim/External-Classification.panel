__title__ = "Set URL"
__doc__ = "Set URL for External Classification API"

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, Entity, AccessLevel
from pyrevit import forms
import clr
clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System.Xaml')
import System
from System.Windows import Window, Application

doc = __revit__.ActiveUIDocument.Document

DEFAULT_URL = "https://api.bsdd.buildingsmart.org"

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

def get_stored_url():
    schema = Schema.Lookup(SCHEMA_GUID)
    if schema:
        entity = doc.ProjectInformation.GetEntity(schema)
        if entity and entity.IsValid():
            field = schema.GetField(FIELD_NAME)
            stored = entity.Get[System.String](field)
            if stored:
                return stored
    return DEFAULT_URL

def set_url():
    xaml_string = """<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Set BSDD URL" Height="220" Width="450"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#F5F5F5">
    <Grid Margin="15">
        <StackPanel VerticalAlignment="Stretch">
            <TextBlock Text="External Classification API URL:" FontSize="12" FontWeight="Bold" Margin="0,0,0,8"/>
            <TextBox Name="UrlTextBox" Height="35" Padding="8" FontSize="11"
                     BorderThickness="1" BorderBrush="#CCCCCC"
                     Background="White" Margin="0,0,0,8"/>
            <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                <Button Name="SaveButton" Content="Save" Width="100" Height="35" 
                        Background="#0078D4" Foreground="White" FontWeight="Bold" Margin="0,0,8,0"/>
                <Button Name="ResetButton" Content="Reset to bSDD" Width="130" Height="35"
                        Background="#F3F3F3" Foreground="Black" BorderThickness="1" BorderBrush="#CCCCCC" Margin="0,0,8,0"/>
                <Button Name="CancelButton" Content="Cancel" Width="100" Height="35"
                        Background="#F3F3F3" Foreground="Black" BorderThickness="1" BorderBrush="#CCCCCC"
                        HorizontalAlignment="Right"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>"""
    
    from System.Windows.Markup import XamlReader
    from System.IO import MemoryStream
    from System.Text import Encoding
    
    try:
        # Convert XAML string to byte stream
        byte_array = Encoding.UTF8.GetBytes(xaml_string)
        stream = MemoryStream(byte_array)
        window = XamlReader.Load(stream)
        url_textbox = window.FindName("UrlTextBox")
        save_button = window.FindName("SaveButton")
        reset_button = window.FindName("ResetButton")
        cancel_button = window.FindName("CancelButton")
        
        # Set initial URL
        url_textbox.Text = get_stored_url()
        
        def save_click(sender, args):
            input_url = url_textbox.Text.strip()
            if input_url:
                schema = get_or_create_schema()
                field = schema.GetField(FIELD_NAME)
                entity = Entity(schema)
                entity.Set[System.String](field, input_url)
                
                with Transaction(doc, "Set External Classification API URL") as t:
                    t.Start()
                    doc.ProjectInformation.SetEntity(entity)
                    t.Commit()
                
                forms.alert("URL set successfully!", title="Success")
                window.Close()
            else:
                forms.alert("Please enter a valid URL.", title="Error")
        
        def reset_click(sender, args):
            url_textbox.Text = DEFAULT_URL
        
        def cancel_click(sender, args):
            window.Close()
        
        save_button.Click += save_click
        reset_button.Click += reset_click
        cancel_button.Click += cancel_click
        
        window.ShowDialog()
    except Exception as e:
        forms.alert("Error opening dialog: {}".format(str(e)), title="Error")

if __name__ == "__main__":
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
            entity.Set[System.String](field, DEFAULT_URL)
            with Transaction(doc, "Set Default External Classification API URL") as t:
                t.Start()
                doc.ProjectInformation.SetEntity(entity)
                t.Commit()
    except Exception:
        pass

    set_url()