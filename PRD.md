### **Product Design / Requirements Document: Rebol 3 Script for Deduplicating Target Folder using `czkawka-CLI.exe`**


**(This document incorporates the change to use a third folder for moving files)**

#### **1. Overview**

* **Product Name:** `dedup-target-folder.r3`
* **Type:** Standalone Rebol 3 script (Oldes Branch - Bulk)
* **Target Platform:** Windows 11
* **Primary Goal:** Automate the process of identifying redundant files within a specified "target" folder by comparing them against a designated "reference" folder, using `czkawka-CLI.exe` for duplicate detection. Identified redundant files from the target folder will be **moved** to a specified third "destination" folder.
* **Integration:** The script relies on the external `czkawka-CLI.exe` tool, which must be present in the system's executable path.

#### **2. Functional Requirements**

The script must replicate the core logic of the provided PowerShell script, adapted for the new "move to folder" strategy:

1. **Input Parameters:**
   * `--reference-folder-path`: Path to the reference folder (mandatory string).
   * `--target-folder-path`: Path to the target folder (mandatory string).
   * `--destination-folder-path`: Path to the destination folder where duplicate target files will be moved (mandatory string).
     * *Note: The script should validate that this folder exists. It should not be the same as the reference or target folder.*
   * `--output-json-file`: Name of the intermediate JSON output file (optional string, default: `duplicates_output.json`).
   * `--dry-run`: Flag to simulate the process without making file system changes (optional switch).
2. **Initial Checks (Intra-Folder Duplicates):**
   * Execute `czkawka-CLI.exe dup` to check the `--target-folder-path` for *internal* duplicates.
   * If duplicates are found within the target folder, print a clear error message and exit the script.
3. **Main Duplicate Detection (Cross-Folder):**
   * Execute `czkawka-CLI.exe dup` with `--directories` set to both `--reference-folder-path` and `--target-folder-path`.
   * Instruct `czkawka` to save the results in compact JSON format to the file specified by `--output-json-file`.
   * If `--dry-run` is active, pass the `--dry-run` flag to `czkawka-CLI.exe`.
4. **Parsing `czkawka` Output:**
   * Read the generated JSON file.
   * Use the `load-json` function (available in Oldes Branch - Bulk) to parse the JSON structure.
   * Navigate the parsed data structure (dictionary of sizes, containing arrays of groups, where each group is an array of file objects) to identify duplicate file groups.
5. **Identifying Target Files:**
   * Iterate through the parsed duplicate groups.
   * For each group, filter the list of files to find those whose `path` resides within the `--target-folder-path`.
6. **Action Based on Mode:**
   * **Dry-Run Mode (`--dry-run`):**
     * Print a list of the identified target folder files that *would be* moved.
     * Exit the script without making any file system changes.
   * **Normal Mode:**
     * Iterate through the list of identified target folder files.
     * For each file, **move** it from its current location in the `--target-folder-path` to the `--destination-folder-path`. Use Rebol's native `rename` function for this.
     * Provide console feedback for each file operation (e.g., "Attempting to move: ...", " -> Success").
     * Handle potential errors during file operations gracefully (e.g., file in use, permission denied, file already exists in destination) and report them.
7. **Final Reporting:**
   * Print a summary upon completion, including:
     * Total number of target files identified for action.
     * Number of files successfully moved (Normal Mode) or that would be moved (Dry-Run Mode).
     * Number of errors encountered.
     * Path to the `czkawka` output JSON file.

#### **3. Non-Functional Requirements**

1. **Usability:**
   * The script should be executable as `dedup-target-folder.r3 --reference-folder-path <path> --target-folder-path <path> --destination-folder-path <path> [--output-json-file <file>] [--dry-run]`.
   * Clear and informative console output, mirroring the PowerShell script's style.
2. **Robustness:**
   * Validate input paths (existence, type - directory).
   * Ensure `--destination-folder-path` is different from `--reference-folder-path` and `--target-folder-path`.
   * Check for the existence and successful execution of `czkawka-CLI.exe`.
   * Gracefully handle errors during file I/O, process execution, and JSON parsing.
   * Handle cases where a file being moved already exists in the destination folder.
3. **Portability & Distribution:**
   * Distributed as a single `.r3` file.
   * Utilizes only standard Rebol 3 (Oldes Branch - Bulk) features and functions (`call/wait/shell/output`, `load-json`, `rename`, file operations, etc.).
4. **Performance:**
   * Efficient parsing of potentially large JSON output from `czkawka`.

#### **4. Technical Design**

1. **Core Components:**
   * **Argument Parser:** To handle command-line arguments.
   * **Process Executor:** Wrapper function for calling `czkawka-CLI.exe` with appropriate arguments and capturing output/exit codes.
   * **JSON Parser & Analyzer:** Function to load the JSON file using `load-json` and navigate its structure to find files within the target folder.
   * **File Mover:** Function to move files using Rebol's `rename` from the target folder to the destination folder. Includes error handling for common issues like existing files.
   * **Main Workflow Controller:** Orchestrates the sequence of checks, calls, parsing, and actions based on the `--dry-run` flag.
2. **Data Flow:**
   * **Input:** Command-line arguments.
   * **Processing:**
     1. Validate arguments.
     2. Call `czkawka` for intra-target-folder check.
     3. If check passes, call `czkawka` for cross-folder check (generating JSON).
     4. Load and parse the JSON file.
     5. Identify target folder files from parsed data.
     6. Depending on `--dry-run`, either list files or attempt to move them to the `--destination-folder-path`.
   * **Output:** Console messages, summary report, potentially modified file system (files moved to the destination folder).
3. **Error Handling:**
   * Use `attempt`, `try`, and specific error checking for file operations, process calls, and JSON parsing.
   * Provide user-friendly error messages for common issues (e.g., `czkawka` not found, file access denied, destination folder issues).

#### **5. Assumptions & Dependencies**

* **`czkawka-CLI.exe`:** Is correctly installed and accessible via the system PATH.
* **Rebol 3 Environment:** The script will run on the Rebol 3 Oldes Branch (Bulk) interpreter, ensuring `load-json` and other necessary functions are available.
* **Destination Folder:** The specified `--destination-folder-path` exists and is writable. It is distinct from the reference and target folders.
* **File Move Strategy:** Standard Rebol `rename` function is used for moving files.

