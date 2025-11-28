# Roadmap Management Tool

Automation tool for managing **CE VHST Roadmaps** streamlining tasks such as time tracking, LC updates, and interface creation management.

---

## 🧩 Features

* **Pointage Automation**: Exports collaborator time tracking data to XML format for VBA import
* **LC Update**: Updates all conditional lists (LC) across template and collaborator files
* **Interface Creation**: Automatically generates user Excel interfaces with three processing modes
* **Interface Deletion or Archiving**: Safely remove or archive all user files with timestamped backups
* **VBA Integration**: Seamless integration with Excel VBA macros for user-friendly workflows
* **Parallel Processing**: Fast interface creation using multiprocessing (9s for 51 files)
* **CLI-based** — fully automatable and compatible with scripts or cron jobs
* **Comprehensive Logging**: All operations logged to `.logs/roadmap.log`

---

## ⚙️ Prerequisites

* **Python 3.11+**
* [**uv**](https://github.com/astral-sh/uv) package manager (recommended)
* **Microsoft Excel** (for VBA integration, optional)

---

## 🧰 Installation

1. Install `uv` (if not installed yet):

   ```bash
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

2. Sync dependencies:

   ```bash
   uv sync
   ```

   This creates a virtual environment and installs required packages automatically.

---

## ▶️ Usage

Run the CLI tool using any of the following methods:

### Option 1 — via `uv`

```bash
uv run roadmap <command> [options]
```

### Option 2 — via virtual environment

```bash
source .venv/bin/activate   # (Windows: .venv\Scripts\activate)
roadmap <command> [options]
```

---

## 🧠 Commands & Options

### Available Functions

#### 0. Base directory

```bash
roadmap --basedir [BASEDIR] <command>
```

#### 1. Pointage (Time Tracking Export)

Exports time tracking data from all collaborator Excel files to an XML file that can be imported by VBA macros.

* Reads data from `POINTAGE` sheet, starting at row 4, columns A-K (11 columns)
* Stops reading when encountering a fully empty row
* Exports to `pointage_output.xml` in the base directory
* Creates empty XML file if no data exists (required for VBA compatibility)
* Skips temporary Excel files (files starting with `~$`)

```bash
roadmap pointage [--delete]
```

**Options:**

* `--delete` → Archive the SYNTHESE sheet instead of exporting pointage data
  * Creates a timestamped archive file containing only SYNTHESE and LC sheets
  * Safely handles open Excel files using temporary file approach

**Examples:**

```bash
# Export pointage data from all collaborator files
roadmap pointage

# Archive SYNTHESE sheet
roadmap pointage --delete
```

**Output:**

Creates `pointage_output.xml` with structure:
```xml
<rows>
  <row>
    <col1>value1</col1>
    <col2>value2</col2>
    ...
  </row>
</rows>
```

#### 2. Update LC (Conditional Lists)

Synchronizes conditional lists (dropdown options) across all Excel files.

* Reads LC data from `Synthèse_RM_CE.xlsm` (columns B-I, starting at row 2)
* Updates the template file (`RM_template.xlsx`)
* Updates all collaborator interface files in `RM_Collaborateurs` folder
* Clears existing LC data before writing new data
* Preserves sheet structure and formulas

```bash
roadmap update
```

**Behavior:**

* Reads LC data from `LC` sheet in `Synthèse_RM_CE.xlsm`
* Updates `RM_template.xlsx` (template file)
* Updates all `RM_[Name].xlsx` files in `RM_Collaborateurs` folder
* For collaborator files, recreates data validation lists in POINTAGE sheet (columns D-G) to ensure consistency
* Preserves existing data in collaborator files - only LC sheet and data validation are updated
* Skips temporary Excel files (files starting with `~$`)

#### 3. Create Interfaces (Creation)

Creates individual Excel interface files for each collaborator listed in the synthesis file.

* Reads collaborator names from `Gestion_Interfaces` sheet, column B (starting at row 3)
* Creates `RM_[COLLABORATOR_NAME].xlsx` files in `RM_Collaborateurs` folder
* Sets collaborator name in cell B1 of POINTAGE sheet
* Adds data validation lists for:
  - Column D: Week (from POINTAGE!A2:A2)
  - Column E: Key (from LC!B3:B1000)
  - Column F: Label (from LC!C3:C1000)
  - Column G: Function (from LC!D3:D1000)

```bash
roadmap create [--way MODE] [--archive]
```

**Options:**

* `--way MODE` → Choose processing mode:
  * `normal` (default): Sequential processing using openpyxl (~50s for 51 files)
  * `para`: Parallel processing using multiprocessing (~9s for 51 files) ⚡ **Fastest**
  * `xlw`: xlwings-based processing (~3min4s for 51 files) - Best for VBA integration
* `--archive` → Archive existing `RM_Collaborateurs` folder before creating new interfaces

**Examples:**

```bash
# Create interfaces using default (normal) mode
roadmap create

# Create interfaces in parallel (fastest)
roadmap create --way para

# Create interfaces and archive existing ones
roadmap create --archive --way para

# Use xlwings mode for better Excel integration
roadmap create --way xlw
```

**Behavior:**

* Reads collaborator names from `Gestion_Interfaces` sheet, column B (starting at row 3)
* Creates `RM_[Name].xlsx` files under `RM_Collaborateurs` folder
* Sets `POINTAGE!B1` to collaborator name
* Adds data validation dropdowns to columns D-G

#### 4. Delete Interfaces

Removes or archives all collaborator interface files.

* Moves `RM_Collaborateurs` folder to `Deleted` directory with timestamp
* Optionally archives to `Archived` directory first
* **Requires `--force` flag** to prevent accidental deletion

```bash
roadmap delete [--archive] --force
```

**Options:**

* `--archive` → Archives files to `Archived` folder before moving to `Deleted`
* `--force` → **Required** to actually perform the operation (safety mechanism)

**Examples:**

```bash
# Archive and delete all files
roadmap delete --archive --force

# Permanently delete all files (no archive)
roadmap delete --force

# Dry run (shows warning, no action taken)
roadmap delete --archive
```

**Behavior:**

* If `--archive` is used: copies folder to `Archived/Archive_RM_Collaborateurs_[timestamp]`
* Moves folder to `Deleted/Deleted_RM_Collaborateurs_[timestamp]`
* Logs count of files processed

---

## 📁 Project Structure

```text
roadmap_manager/
│   .gitignore
│   pyproject.toml              # Project configuration & dependencies
│   README.md                   # This file
│   uv.lock                     # Dependency lock file
│   VBA_code.bas                # Excel VBA macros for integration
│   
├───.logs/
│       roadmap.log             # Application logs
│
├───roadmap/                    # Main Python package
│       __init__.py             # Package initialization
│       main.py                 # RoadmapManager class & CLI entry point
│       helpers.py              # Utility functions (XML, parsing, validation)
│
└───tests/                      # Unit tests
        test_cli.py             # CLI argument parsing tests
        test_helpers.py         # Helper function tests
```

## 📂 Expected File Structure (Base Directory)

The tool expects the following structure in your base directory:

```text
base_directory/
│   Synthèse_RM_CE.xlsm         # Master synthesis file (required)
│   RM_template.xlsx             # Template file for interfaces (required)
│   pointage_output.xml          # Generated XML export (created by tool)
│
├───RM_Collaborateurs/           # Collaborator interface files (created by tool)
│       RM_Alice.xlsx
│       RM_Bob.xlsx
│       ...
│
├───Archived/                    # Archive folder (created by tool)
│       Archive_RM_Collaborateurs_01012024_120000/
│       Archive_SYNTHESE_01012024_120000.xlsx
│       ...
│
└───Deleted/                      # Deleted files folder (created by tool)
        Deleted_RM_Collaborateurs_01012024_120000/
        ...
```

---

## 🔌 VBA Integration

The tool integrates seamlessly with Excel VBA macros. The `VBA_code.bas` file contains macros that can be imported into your Excel workbook.

### VBA Functions Available:

* `Btn_Create_RM()` - Create interfaces via button click
* `Btn_Delete_RM()` - Delete interfaces with confirmation dialogs
* `Btn_Collect_RM_Data()` - Import pointage data from XML to SYNTHESE sheet
* `Btn_Clear_Synthese()` - Clear SYNTHESE sheet with archiving option
* `Btn_Update_LC()` - Update conditional lists (LC) in template and all collaborator files

### Setup:

1. Import `VBA_code.bas` into your Excel workbook
2. Update `PYTHONEXE` constant in VBA code with your Python executable path
3. Create buttons/controls linked to the VBA functions
4. Ensure the base directory path is set correctly

The VBA code calls the Python CLI tool and handles Excel-specific operations like importing XML data into worksheets.

---

## 📊 Logging

All operations are logged to `.logs/roadmap.log` with the following information:

* Operation start/completion
* File counts and processing status
* Errors and warnings
* Performance metrics

Log format: `%(asctime)s [%(levelname)s] %(message)s`

---

## 🏁 Example Workflows

### Complete Setup Workflow

```bash
# 1. Update conditional lists in all files
roadmap --basedir /path/to/files update

# 2. Create interfaces for all collaborators (parallel mode)
roadmap --basedir /path/to/files create --way para --archive

# 3. Export pointage data
roadmap --basedir /path/to/files pointage
```

### Regular Maintenance Workflow

```bash
# Weekly: Export time tracking data
roadmap pointage

# Monthly: Update conditional lists
roadmap update

# Quarterly: Refresh all interfaces
roadmap create --way para --archive
```

### Cleanup Workflow

```bash
# Archive and delete old interfaces
roadmap delete --archive --force

# Archive SYNTHESE sheet
roadmap pointage --delete
```

---

## 🐛 Troubleshooting

### Common Issues:

**"Required files missing" error:**
- Ensure `Synthèse_RM_CE.xlsm` and `RM_template.xlsx` exist in base directory
- Check file names match exactly (case-sensitive)

**"Template file is opened" error:**
- Close the template Excel file before running create operations

**"No collaborators found" error:**
- Check `Gestion_Interfaces` sheet exists in synthesis file
- Verify collaborator names are in column B starting at row 3

**VBA integration not working:**
- Verify Python executable path in `PYTHONEXE` constant
- Ensure base directory path is correctly set
- Check that `pointage_output.xml` is generated after pointage command

---

## 📝 Notes

* The tool uses temporary files to safely handle open Excel files
* Parallel processing (`--way para`) requires the template file to be closed
* XML export format is designed specifically for VBA parsing
* All archive operations create timestamped folders/files
* The tool skips temporary Excel files (starting with `~$`)
