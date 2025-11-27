# Roadmap Management Tool

Automation tool for managing **CE VHST Roadmaps** streamlining tasks such as time tracking, LC updates, and interface creation management.

---

## 🧩 Features

* **Pointage Automation**: Exports collaborator time tracking data to the synthesis file.
* **LC Update**: Updates all conditional lists (LC) across template and collaborator files.
* **Interface Creation**: Automatically generates user Excel interfaces.
* **Interface Deletion or Archiving**: Safely remove or archive all user files.
* **Header Verification** and **Data Integrity Checks** ensure accuracy.
* **CLI-based** — fully automatable and compatible with scripts or cron jobs.
---

## ⚙️ Prerequisites

* **Python 3.14+**
* [**uv**](https://github.com/astral-sh/uv) package manager (recommended)

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

* Exports time tracking data from collaborator files to the central synthesis
* Copies data from range A4:KX in the POINTAGE sheet
* Appends data to SYNTHESE sheet in Synthèse_RM_CE_VHST.xlsx
* Clears the POINTAGE sheet after export
* Includes header verification

```bash
roadmap pointage [--choice N]
```

**Options:**

* `--choice <int>` → Select which file to process:

  * `-1` (default): process all collaborator files.
  * Any positive integer: process the file at that index.

**Example:**

```bash
roadmap pointage
roadmap pointage --choice 3
```

#### 2. Update LC (Conditional Lists)

* Updates conditional lists across all personal tools
* Copies LC data (B3:IX) from Synthèse_RM_CE_VHST.xlsx
* Updates the template file (RM_NOM Prénom.xlsx)
* Updates all files in RM_Collaborateurs folder
* Preserves sheet structure and formulas

```bash
roadmap update
```

**Behavior:**

* Copies LC data from **Synthèse_RM_CE_VHST.xlsx**
* Updates both **RM_NOM Prénom.xlsx** (template) and all files under `RM_Collaborateurs`

#### 3. Create Interfaces (Creation)

* Creates personal tools for all collaborators listed in LC sheet
* Reads collaborator names from column B (starting row 3) in LC sheet
* Duplicates template file for each missing collaborator
* Names files as: RM_[COLLABORATOR_NAME].xlsx
* Sets collaborator name in cell B1 of POINTAGE sheet
* Protects POINTAGE sheet

```bash
roadmap create
```

**Behavior:**

* Reads collaborator names from column **B** in LC sheet (starting at row 3)
* Creates missing `RM_[Name].xlsx` files under `RM_Collaborateurs`
* Sets `POINTAGE!B1` to collaborator name and enables sheet protection

#### 4. Delete Interfaces

* Removes all files from RM_Collaborateurs folder
* Option to archive (rename with "Archive_" prefix) instead of deleting
* Confirmation required before execution

```bash
roadmap delete --archive [--force]
```

**Options:**

* `--archive` → Archives files (renames with `Archive_` prefix)
* no `--archive` → Deletes files permanently
* `--force` → Required to actually perform the operation (safety mechanism)

**Examples:**

```bash
# Archive all files
roadmap delete --archive --force

# Permanently delete all files
roadmap delete --force

# Dry run (no --force)
roadmap delete --archive
```

---

## 📁 Folder Structure

```text
project_root/
│   .gitignore
│   .python-version
│   pyproject.toml
│   README.md
│   uv.lock
│   
├───.logs
│       roadmap.log
│
├───roadmap
│       helpers.py
│       main.py                  # CLI entry point
│       __init__.py
│
├───scripts
│       data_validation_list.py
│       rm.xlsx
│       test.py
│
└───tests
        test_cli.py
        test_helpers.py
```

---

## 🏁 Example Workflow

```bash
# Export pointage data from all collaborator files
roadmap pointage

# Update LC data in all files
roadmap update

# Create missing user interfaces
roadmap create

# Delete or archive all collaborator files (requires --force)
roadmap delete --archive --force
```
