# External Classification

External Classification is a pyRevit extension for Revit 2024 and newer. It lets you assign classifications and properties from bSDD (or other Data Dictionaries with bSDD REST API structure) to elements and export them into IFC using IfcClassification and IfcClassificationReference.

## Requirements

- Revit 2024 or newer
- pyRevit 4.8 or newer
- Internet access for bSDD lookup

## Notes

- The extension is tested in Revit 2024 and 2025.
- Up to five different dictionaries can be used in one project.
- The export workflow is intended for IFC-based exports of version 4 or higher. However, it also works with version 2x3, though not all dictionary information can be exported to IFC.

## Installation

### Step 1: Install pyRevit

If pyRevit is not already installed, install it first.

Recommended options:

1. Download the latest [pyRevit](https://pyrevitlabs.notion.site/Install-pyRevit-98ca4359920a42c3af5c12a7c99a196d) installer from the official [pyRevit release page](https://github.com/pyrevitlabs/pyRevit/releases).
2. Run the installer and follow the setup instructions.
3. Start Revit and confirm that a pyRevit tab appears in the ribbon.

If you already have pyRevit installed, you can skip this step.

### Recommended: install from GitHub

1. Download or clone this repository from GitHub.
2. Rename the unziped folder to `External Classification.panel` and copy it into your pyRevit extensions folder.
3. Restart Revit.

Typical pyRevit extensions folder locations in Windows are:

- `C:\Users\[username]\%APPDATA%\Roaming\pyRevit\Extensions`
- `C:\Users\[username]\%APPDATA%\Roaming\pyRevit\Custom Extensions`

If you already use pyRevit, place the panel in the same extensions location where your other custom tools are stored.

If you are not sure whether pyRevit is installed, open Revit and check whether a pyRevit tab is visible. If the tab is missing, install pyRevit first.

### Manual install from a ZIP file

1. Open the repository on GitHub.
2. Download the ZIP archive.
3. Extract the archive.
4. Copy the `External Classification.panel` folder into your pyRevit extensions folder.
5. Restart Revit.

## First Start

After installation, the extension appears in the pyRevit ribbon.

### Start in 5 steps

1. Open Revit.
2. Open the pyRevit Bundles Creator tab.
3. Open `Set URL` if you want to use a custom bSDD endpoint.
4. Use `Add Classification` to assign classes to selected elements.
5. Use `Export IFC` to include classification data in the IFC file.
6. 
If there are any problems, see the section for troubleshooting.

### Command Overview

- `Set URL`: saves the API endpoint in the project.
- `Add Classification`: lets you pick a dictionary and class, then apply it to one or many elements.
- `Export IFC`: exports the model and injects classification data into the IFC output.

### Classification Workflows

- **Direct class assignment**: assign a selected class to the current selection.
- **Dictionary-first selection**: choose the dictionary first, then the class.
- **Single and multi-element assignment**: classify one element or a batch at once.
- **Multiple classifications per element**: add classifications from different dictionaries to the same element.
- **Multiple dictionaries per project**: use several dictionaries in one model.

### Property Workflows

- **Add class-based properties**: if a class contains properties, fill them in during classification.
- **Per-element or shared values**: decide whether values are stored per element or for multiple elements together.
- **Skip properties when needed**: clear `Open properties after classification` to classify without opening the property dialog.
- **Edit existing values**: update property values later in the Properties tab.
- **Delete classifications**: remove classifications in the Overview tab.

### IFC Export Options

- **Use Revit IFC export profiles**: choose an existing export profile.
- **Flat or hierarchical classification**:
	- **Flat**: export only the selected class.
	- **Full**: export the selected class and its parent hierarchy.

## Technical Implementation

This extension combines Revit Extensible Storage, shared parameters, and IFC export metadata to keep classification data stable in the model and transferable to IFC.

### Extensible Storage (internal model data)

- **Project-level API URL** (`DataDictionaryAPI` schema): stores the configured bSDD endpoint on `ProjectInformation`.
- **Element-level classification** (`ExternalClassification` schema): stores a legacy single classification payload (`Code`, `Name`, `ClassUri`, `DictionaryName`, `DictionaryUri`) for backward compatibility.
- **Element-level multi-classification** (`ExternalClassificationMulti` schema): stores multiple classification items as JSON (`ItemsJson`) per element.
- **Element-level class properties** (`ExternalClassificationProperties` schema): stores class property sets as JSON (`PropertiesJson`) per element and per class URI.

### Shared Parameters

- The extension uses its own shared parameter definition file (`02_AddClassification.pushbutton/external_shared_params.txt`) for classification-related parameters.
- During IFC export, built-in IFC shared parameters are resolved from Revit installation paths (for example `IFC Shared Parameters-RevitIFCBuiltIn_ALL.txt`) and bound to relevant categories when needed.

### IFC Mapping Strategy

- For IFC export, the tool prepares IFC classification metadata via an `IFCClassification` Extensible Storage schema on `DataStorage` elements.
- These entries are used to map dictionary/system data to IFC `IfcClassification` and `IfcClassificationReference`.
- Depending on export mode, references are written as either flat class links or hierarchical class chains (class plus ancestors).

## Troubleshooting

- If the extension does not appear, restart Revit and check that the folder was copied to the pyRevit extensions directory.
- If the click on a button throws an exception, close Revit, go to `C:\Users\[username]\AppData\Roaming\pyRevit\[version]` and delete the `pyRevit_version_..._pyRevitBundlesCreatorExtension.dll`. Then restart Revit.
- If bSDD requests fail, check your internet connection or the configured API URL.
- This extension is only intended for Revit 2024 and newer.

## AI Declaration
This project was developed with AI-assisted tools, including GitHub Copilot and Claude Sonnet 4.6.

## License

See [LICENSE](LICENSE).