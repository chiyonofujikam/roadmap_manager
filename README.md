# Roadmap Management Tool

Automation tool for managing **CE Roadmaps** streamlining tasks such as time tracking, LC updates, and interface creation management.

---

## Features

* **Pointage Automation**: Exports collaborator time tracking data to XML format for VBA import
* **LC Update**: Updates all conditional lists (LC) across template and collaborator files
* **Interface Creation**: Automatically generates user Excel interfaces with three processing modes
* **Interface Deletion or Archiving**: Safely remove or archive all user files with timestamped backups
* **Cleanup Missing Collaborators**: Automatically removes interface files for collaborators no longer in the list
* **VBA Integration**: Seamless integration with Excel VBA macros for user-friendly workflows
* **Parallel Processing**: Fast interface creation using multiprocessing (~9s for 51 files)
* **CLI-based**: Fully automatable and compatible with scripts or scheduled tasks
* **Comprehensive Logging**: All operations logged to `.logs/roadmap.log`
* **Executable Build**: Can be packaged as standalone `.exe` for distribution

---

## Prerequisites

* **Python 3.11+**
* [**uv**](https://github.com/astral-sh/uv) package manager
* **Microsoft Excel** (for VBA integration and file operations)
* **Windows OS** (primary platform, though code may work on other platforms)

---

## Installation

1. Install `uv` (if not installed yet):

   ```bash
   curl -LsSf https://astral.sh/uv/install.sh | sh
   # Windows PowerShell:
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   ```

2. Clone or navigate to the project directory:

   ```bash
   cd roadmap_manager
   ```

3. Sync dependencies:

   ```bash
   uv sync
   .venv\Scripts\activate
   ```

   This creates a virtual environment and installs required packages automatically.

---

## Usage

Run the CLI tool using any of the following methods:

```bash
uv run roadmap <command> [options]
```
---

## Commands & Options

### Global Options

#### Base Directory

```bash
roadmap --basedir [BASEDIR] <command>
```

Specify the base directory path containing roadmap files

---

### Available Commands

#### 1. Pointage (Time Tracking Export)

Exports time tracking data from all collaborator Excel files to an XML file that can be imported by VBA macros.

**What it does:**
* Reads data from `POINTAGE` sheet, starting at row 4, columns A-K (11 columns)
* Stops reading when encountering a fully empty row
* Exports to `pointage_output.xml` in the base directory
* Creates empty XML file if no data exists (required for VBA compatibility)
* Skips temporary Excel files (files starting with `~$`)

```bash
roadmap pointage
```

**Output:**

Creates `pointage_output.xml` with structure:
```xml
<?xml version='1.0' encoding='utf-8'?>
<rows>
  <row>
    <col1>value1</col1>
    <col2>value2</col2>
    <col3>value3</col3>
    ...
    <col11>value11</col11>
  </row>
  ...
</rows>
```

**Examples:**

```bash
# Export pointage data from all collaborator files
roadmap pointage

# With custom base directory
roadmap --basedir "C:\MyRoadmapFiles" pointage
```

---

#### 2. Update LC (Conditional Lists)

**Note:** LC Update is now **fully implemented in VBA** and no longer requires Python CLI. The `roadmap update` command is kept for backward compatibility but the VBA button (`Btn_Update_LC`) performs all operations directly in Excel.

**VBA Implementation (`Btn_Update_LC`):**

**What it does:**
* Reads LC data directly from `LC` sheet in `Synthèse_RM_CE.xlsm` (no temporary file needed)
* Reads columns B-I (columns 2-9), starting at row 2
* Updates the template file (`RM_template.xlsx`)
* Updates all collaborator interface files in `RM_Collaborateurs` folder
* Preserves cell values exactly (text format to prevent date/number conversion)
* Processes files efficiently with progress tracking
* Shows completion time in the success message
* Windows remain hidden during processing (no visual flashing)
* Ensures windows are visible before closing (prevents saving hidden state)

**Performance:**
* Optimized bulk operations using Copy/PasteSpecial
* Calculation disabled during processing for speed
* Processes 20+ files efficiently
* Shows progress in status bar: "Updating LC: X of Y files..."

**Usage:**
* Click "Update LC" button in Excel (VBA macro)
* Confirms operation before proceeding
* Shows completion message with file count and time taken

**Python CLI (Legacy):**

The `roadmap update` command is still available for backward compatibility but is deprecated in favor of the VBA implementation.

```bash
roadmap update
```

**Prerequisites:**
* Template and collaborator files must not be open in Excel
* `LC` sheet must exist in `Synthèse_RM_CE.xlsm`

---

#### 3. Create Interfaces

Creates individual Excel interface files for each collaborator listed in the synthesis file.

**What it does:**
* Reads collaborator names from `collabs.xml` file (created by VBA macros)
* Creates `RM_[COLLABORATOR_NAME].xlsx` files in `RM_Collaborateurs` folder
* Sets collaborator name in cell B1 of POINTAGE sheet
* Adds data validation lists for:
  - Column D: Week (from POINTAGE!A2:A2)
  - Column E: Key (from LC!B3:B1000)
  - Column F: Label (from LC!C3:C1000)
  - Column G: Function (from LC!D3:D1000)
* Only creates files that don't already exist (skips existing files)

```bash
roadmap create [--way MODE] [--archive]
```

**Options:**

* `--way MODE` → Choose processing mode:
  * `normal` (default): Sequential processing using openpyxl (~50s for 51 files)
  * `para`: Parallel processing using multiprocessing (~9s for 51 files) **Fastest**
* `--archive` → Archive existing `RM_Collaborateurs` folder before creating new interfaces

**Examples:**

```bash
# Create interfaces using default (normal) mode
roadmap create

# Create interfaces in parallel (fastest)
roadmap create --way para

# Create interfaces and archive existing ones
roadmap create --archive --way para

# With custom base directory
roadmap --basedir "C:\MyRoadmapFiles" create --way para
```

**Prerequisites:**
* `collabs.xml` file must exist in the base directory (created by VBA macros)
* Template file (`RM_template.xlsx`) must be closed
* For parallel mode, template file must be accessible (not locked)

**Note:** The `collabs.xml` file is automatically deleted after reading to keep the directory clean.

---

#### 4. Delete Interfaces

Removes or archives all collaborator interface files.

**What it does:**
* Moves `RM_Collaborateurs` folder to `Deleted` directory with timestamp
* Optionally archives to `Archived` directory first
* Creates zip archives with timestamped names
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

**Archive Structure:**
* If `--archive` is used: `Archived/Archive_RM_Collaborateurs_[timestamp].zip`
* Always creates: `Deleted/Deleted_RM_Collaborateurs_[timestamp].zip`
* Timestamp format: `DDMMYYYY_HHMMSS`

---

#### 5. Cleanup Missing Collaborators

Deletes interface files for collaborators that are missing from the current collaborator list.

**What it does:**
* Compares existing files in `RM_Collaborateurs` folder with collaborator list from `collabs.xml`
* If a file exists but the collaborator is not in the XML list, that file is deleted
* Creates a zip archive of deleted files before deletion
* Skips temporary Excel files (files starting with `~$`)

```bash
roadmap cleanup
```

**Examples:**

```bash
# Remove interfaces for missing collaborators
roadmap cleanup

# With custom base directory
roadmap --basedir "C:\MyRoadmapFiles" cleanup
```

**Prerequisites:**
* `collabs.xml` file must exist in the base directory (created by VBA macros)

**Archive:**
* Creates `Deleted/Deleted_Missing_RM_collaborators_[timestamp].zip` before deletion

---

## Documentation

The project includes comprehensive documentation:

| Document | Description |
|----------|-------------|
| `README.md` | Main documentation (this file) - installation, usage, commands |
| `PRESENTATION.md` | Project presentation with Mermaid diagrams - architecture overview, workflows |
---

## Project Structure

```text
roadmap_manager/
│   .gitignore
│   pyproject.toml              # Project configuration & dependencies
│   README.md                   # This file
│   PRESENTATION.md             # Project presentation with diagrams
│   uv.lock                     # Dependency lock file (uv)
│   roadmap_cli.py              # Entry point for PyInstaller builds
│   build_exe.bat               # Batch script to build executable
│
├───.logs/                      # Log directory (created automatically)
│       roadmap.log             # Application logs
│
├───roadmap/                    # Main Python package
│       __init__.py             # Package initialization
│       main.py                 # CLI entry point and argument parsing
│       roadmap.py              # RoadmapManager class (core logic)
│       helpers.py              # Utility functions (XML, parsing, validation)
│
├───VBA/                        # VBA integration code
│       modButtonHandlers.bas   # Button click event handlers
│       modGlobals.bas          # Global constants and variables
│       modUtilities.bas        # Utility functions for VBA
│
├───tests/                      # Unit & integration tests
│       __init__.py             # Test package initialization
│       conftest.py             # Shared pytest fixtures
│       test_cli.py             # CLI argument parsing tests
│       test_helpers.py         # Helper function tests
│       test_roadmap_manager.py # RoadmapManager integration tests
│
├───htmlcov/                    # Coverage report (generated, gitignored)
│
├───build/                      # PyInstaller build artifacts (temporary)
│
└───dist/                       # PyInstaller output directory
        roadmap.exe             # Built executable (if built)
```

---

## Expected File Structure (Base Directory)

The tool expects the following structure in your base directory:

```text
base_directory/
│   Synthèse_RM_CE.xlsm         # Master synthesis file (required)
│   RM_template.xlsx             # Template file for interfaces (required)
│   collabs.xml                  # Temporary file (created by VBA, deleted after use)
│   pointage_output.xml          # Generated XML export (created by tool)
│
├───script/                      # Executable location (for VBA integration)
│       roadmap.exe              # Built executable (copied here for VBA)
│
├───RM_Collaborateurs/           # Collaborator interface files (created by tool)
│       ...
│
├───Archived/                    # Archive folder (created by tool)
│       Archive_RM_Collaborateurs_01012025_120000.zip
│       Archive_SYNTHESE_01012024_120000.xlsx    # Contains SYNTHESE + LC sheets
│       ...
│
└───Deleted/                      # Deleted files folder (created by tool)
        Deleted_RM_Collaborateurs_01012024_120000.zip
        Deleted_Missing_RM_collaborators_01012024_120000.zip
        ...
```

**Note:** The VBA code expects `roadmap.exe` in the `script/` subdirectory of the base directory.

### Required Excel File Structure

**Synthèse_RM_CE.xlsm:**
* Must contain `Gestion_Interfaces` sheet with collaborator names in column B (starting at row 3)
* Must contain `LC` sheet with conditional list data (columns B-I, starting at row 2)
* Must contain `SYNTHESE` sheet for pointage data import

**RM_template.xlsx:**
* Must contain `POINTAGE` sheet with:
  - Cell A2: Week value (for dropdown)
  - Cell B1: Will be set to collaborator name
  - Columns D-G: Will have data validation lists added
* Must contain `LC` sheet with conditional list structure

---

## VBA Integration

The tool integrates seamlessly with Excel VBA macros. The `VBA/` directory contains VBA modules that can be imported into your Excel workbook.

### VBA Modules

1. **modGlobals.bas**: Global constants and variables
   * `PYTHONEXE`: Path to Python executable or roadmap.exe
   * `GetBaseDir()`: Function to get base directory path

2. **modUtilities.bas**: Utility functions
   * `RunCommand()`: Execute shell commands and return exit code
   * `GetBaseDir()`: Get or prompt for base directory path (cached)
   * `LoadXMLTable()`: Parse XML file and return data as collection
   * `EscapeXML()`: Escape special characters for XML content
   * `CreateCollabsXML()`: Generate collaborator XML from Gestion_Interfaces sheet
   * `CreateLCExcel()`: Legacy function (deprecated - LC update now VBA-only)
   * `CleanupGestionInterfaces()`: Remove empty rows from interface sheet (auto-runs before operations)

3. **modButtonHandlers.bas**: Button click event handlers
   * `Btn_Create_RM()`: Create interfaces via button click
   * `Btn_Delete_RM()`: Delete interfaces with confirmation dialogs
   * `Btn_Collect_RM_Data()`: Import pointage data from XML to SYNTHESE sheet
   * `Btn_Collect_RM_Data_Reset()`: Full reset workflow - import data, delete interfaces, and recreate
   * `Btn_Clear_Synthese()`: Archive SYNTHESE and LC sheets, then clear SYNTHESE data
   * `Btn_Update_LC()`: **VBA-only** update of conditional lists (LC) in template and all collaborator files (no Python call)
   * `Btn_Cleanup_RM()`: Cleanup interfaces for collaborators no longer in the list
   * `FixHiddenWindows()`: Utility function to restore window visibility for affected files

### VBA Setup

1. **Import VBA modules:**
   * Open Excel workbook (`Synthèse_RM_CE.xlsm`)
   * Press `Alt + F11` to open VBA Editor
   * Right-click on your project → Import File
   * Import all `.bas` files from the `VBA/` directory, or import `VBA_code.bas` (combined version)

2. **Create buttons/controls:**
   * In Excel, go to Developer tab → Insert → Button (Form Control)
   * Assign macros to buttons:
     * `Btn_Create_RM` → "Create Interfaces"
     * `Btn_Delete_RM` → "Delete Interfaces"
     * `Btn_Collect_RM_Data` → "Collect Pointage Data"
     * `Btn_Collect_RM_Data_Reset` → "Collect & Reset" (full cycle: collect → delete → recreate)
     * `Btn_Clear_Synthese` → "Clear SYNTHESE"
     * `Btn_Update_LC` → "Update LC"
     * `Btn_Cleanup_RM` → "Cleanup Missing"

3. **Verify base directory:**
   * Ensure `GetBaseDir()` function returns the correct path
   * Or modify it to match your file structure

### How VBA Integration Works

1. **VBA prepares data:**
   * `CleanupGestionInterfaces()` removes empty rows from collaborator list (runs automatically)
   * `collabs.xml`: Generated from `Gestion_Interfaces` sheet column B (for create/cleanup operations)

2. **VBA operations:**
   * **LC Update (`Btn_Update_LC`)**: Fully implemented in VBA - reads LC sheet directly and updates all files
   * **Other operations**: VBA calls Python CLI via shell commands

3. **VBA calls Python CLI** (for create, delete, pointage, cleanup):
   * Executes shell command with appropriate arguments via `RunCommand()`
   * Waits for completion and checks exit code

4. **Python processes files** (for create, delete, pointage, cleanup):
   * Reads temporary files (`collabs.xml` for create/cleanup, `pointage_output.xml` for pointage)
   * Performs operations (create, delete, pointage, cleanup)
   * Generates output files (e.g., `pointage_output.xml`)
   * Deletes temporary input files after reading

5. **VBA imports results:**
   * `LoadXMLTable()` parses `pointage_output.xml`
   * Imports data to `SYNTHESE` sheet starting at first empty row
   * Cleans up temporary XML file after import

---

## Building Executable

You can build a standalone executable for distribution without requiring Python installation.

### Prerequisites

* PyInstaller installed: `uv add pyinstaller`
* All dependencies installed

### Build Process

**Using build script (Windows)**

```bash
build_exe.bat
```

This script:
* Builds `roadmap.exe` using PyInstaller
* Copies executable to `dist/roadmap.exe`
* Optionally copies to a predefined destination directory

**Output:**
* Executable: `dist/roadmap.exe`
* Size: ~50-100 MB (includes Python runtime and all dependencies)

**Usage:**
```bash
# Run executable directly
roadmap.exe create --way para

# Or with full path
C:\path\to\roadmap.exe pointage
```

**Note:** The executable is self-contained and doesn't require Python or any dependencies to be installed on the target machine.

---

## Testing

The project includes a comprehensive test suite using pytest.

### Running Tests

```bash
# Run all tests
uv run pytest tests/ -v

# Run specific test file
uv run pytest tests/test_cli.py -v
uv run pytest tests/test_helpers.py -v
uv run pytest tests/test_roadmap_manager.py -v
```

### Code Coverage

Generate a coverage report to see how much of the code is tested:

```bash
# Run tests with coverage report (terminal output)
uv run pytest tests/ --cov=roadmap

# Run tests with HTML coverage report
uv run pytest tests/ --cov=roadmap --cov-report=html
```

The HTML report is generated in the `htmlcov/` directory. Open `htmlcov/index.html` in a browser to view detailed coverage information.

### Test Structure

| Test File | Description | Tests |
|-----------|-------------|-------|
| `test_cli.py` | CLI argument parsing tests | 2 |
| `test_helpers.py` | Helper function tests (XML, Excel, validation) | 26 |
| `test_roadmap_manager.py` | RoadmapManager integration tests | 30 |
| `conftest.py` | Shared pytest fixtures | - |

**Total: 58 tests**

### Test Categories

* **CLI Tests**: Validate command-line argument parsing for all commands
* **Helper Tests**: Test utility functions (XML operations, file handling, data validation)
* **Integration Tests**: Test RoadmapManager class operations (create, delete, pointage, update)

---

## Logging

All operations are logged to `.logs/roadmap.log` with the following information:

* Operation start/completion
* File counts and processing status
* Errors and warnings
* Performance metrics
* Detailed error traces

**Log format:** `%(asctime)s [%(levelname)s] %(message)s`

**Log location:**
* If running as script: `.logs/roadmap.log` in project directory
* If running as executable: `.logs/roadmap.log` next to executable
* Fallback: System temp directory if write permissions unavailable

**Log levels:**
* `DEBUG`: Detailed diagnostic information
* `INFO`: General informational messages
* `WARNING`: Warning messages (non-critical issues)
* `ERROR`: Error messages (operations failed)

---

## Example Workflows

### Complete Setup Workflow

```bash
# 1. Update conditional lists in all files
roadmap --basedir "C:\MyRoadmapFiles" update

# 2. Create interfaces for all collaborators (parallel mode)
roadmap --basedir "C:\MyRoadmapFiles" create --way para --archive

# 3. Export pointage data
roadmap --basedir "C:\MyRoadmapFiles" pointage
```

### Regular Maintenance Workflow

```bash
# Weekly: Export time tracking data
roadmap pointage

# Monthly: Update conditional lists
roadmap update

# Quarterly: Refresh all interfaces
roadmap create --way para --archive

# Cleanup: Remove interfaces for missing collaborators
roadmap cleanup
```

### Cleanup Workflow

```bash
# Archive and delete old interfaces
roadmap delete --archive --force

# Remove interfaces for collaborators no longer in list
roadmap cleanup
```

### VBA-Driven Workflow

1. Open `Synthèse_RM_CE.xlsm` in Excel
2. Click "Update LC" button → Updates all conditional lists
3. Click "Create Interfaces" button → Creates all collaborator files
4. Collaborators fill their time tracking in their individual files
5. Click "Collect Pointage Data" button → Exports and imports all pointage data
6. Click "Clear SYNTHESE" button → Archives and clears synthesis sheet

---

## Troubleshooting

### Common Issues

**"Required files missing" error:**
- Ensure `Synthèse_RM_CE.xlsm` and `RM_template.xlsx` exist in base directory
- Check file names match exactly
- Verify base directory path is correct

**"Template file is opened" error:**
- Close the template Excel file before running create operations
- Check if Excel process is still running in background
- For parallel mode, ensure template file is completely closed

**"No collaborators found" error:**
- Check `collabs.xml` file exists in base directory (created by VBA)
- Verify `Gestion_Interfaces` sheet exists in synthesis file
- Ensure collaborator names are in column B starting at row 3
- Check that VBA macro `CreateCollabsXML()` ran successfully

**LC update errors:**
- Ensure `LC` sheet exists in synthesis file (`Synthèse_RM_CE.xlsm`)
- Verify base directory path is correct
- Check that template and collaborator files are not open in Excel

**VBA integration not working:**
- Verify Python executable path in `PYTHONEXE` constant
- Ensure base directory path is correctly set in `GetBaseDir()`
- Check that `pointage_output.xml` is generated after pointage command
- Verify VBA macros are enabled in Excel (File → Options → Trust Center → Macro Settings)
- Check Windows security settings aren't blocking script execution

**Permission errors:**
- Ensure Excel files are closed before running operations
- Check file/folder permissions in base directory
- Verify you have write access to base directory and subdirectories
- On OneDrive: Ensure files are synced and not in "Files On-Demand" mode

**Parallel processing issues:**
- Ensure template file is closed
- Check available system memory (parallel mode uses more RAM)
- Reduce `max_workers` parameter if system is slow
- Fall back to `normal` mode if issues persist

**Log file not created:**
- Check write permissions in project directory or executable directory
- Verify disk space is available
- Check if antivirus is blocking file creation

---

## Technical Details

### Processing Modes Comparison

| Mode | Library | Speed (51 files) | Use Case |
|------|---------|------------------|----------|
| `normal` | openpyxl | ~50s | Standard use, reliable |
| `para` | openpyxl (multiprocessing) | ~9s | Fast batch creation |
 
### File Handling

* **Temporary files**: `collabs.xml` is automatically deleted after use
* **File locking**: Tool uses temporary file approach to handle open Excel files
* **Retry logic**: Folder deletion includes retry mechanism for Windows/OneDrive locks
* **Skip patterns**: Automatically skips temporary Excel files (starting with `~$`)

### Data Validation

The tool creates Excel data validation lists for:
* **Column D**: Week selection (from POINTAGE!A2:A2)
* **Column E**: Key selection (from LC!B3:B1000)
* **Column F**: Label selection (from LC!C3:C1000)
* **Column G**: Function selection (from LC!D3:D1000)

### XML Format

Pointage XML follows this structure:
```xml
<?xml version='1.0' encoding='utf-8'?>
<rows>
  <row>
    <col1>Value for column A</col1>
    <col2>Value for column B</col2>
    ...
    <col11>Value for column K</col11>
  </row>
</rows>
```

This format is optimized for VBA parsing using `MSXML2.DOMDocument`.

---

## Security Notes

* The tool reads and writes Excel files - ensure files are from trusted sources
* VBA macros require macro-enabled workbooks (`.xlsm` files)
* Executable files should be from trusted sources
* Log files may contain file paths and collaborator names - protect log files appropriately

---

## Author

**Mustapha ELKAMILI**

---

## Dependencies

* **openpyxl** (>=3.1.2): Excel file manipulation
* **tqdm** (>=4.67.1): Progress bars
* **pywin32** (>=306): Windows-specific functionality
* **pyinstaller** (>=6.17.0): Executable building (optional)
* **pytest** (>=9.0.1): Testing framework (dev dependency)
* **pytest-cov** (>=7.0.0): Coverage reporting (dev dependency)

---
