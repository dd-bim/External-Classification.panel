# Development

This project is a pyRevit extension for Revit 2024 and newer.

## Key files

- `revit_compat.py` handles version checks and shared parameter paths.
- `shared_text_utils.py` contains string and IFC escaping helpers.
- `01_SetURL.pushbutton` stores the API endpoint used by the extension.
- `02_AddClassification.pushbutton` manages classification and properties.
- `03_ExportIFC.pushbutton` handles IFC export.

## Local setup

1. Clone the repository.
2. Copy or link `External Classification.panel` into the pyRevit extensions folder.
3. Start Revit 2024 or newer.

