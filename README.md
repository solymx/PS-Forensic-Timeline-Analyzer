


# 🕵️‍♂️ PS Forensic Timeline & MOTW Analyzer

A lightweight, standalone **PowerShell GUI tool** designed for digital forensics and incident response (DFIR). This tool assists analysts in filtering files based on timestamps, analyzing file ownership, and inspecting "Mark of the Web" (MOTW) data to trace the origin of downloaded files.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Version](https://img.shields.io/badge/version-v8.0-green)

## 🚀 Key Features

* **GUI Interface:** User-friendly Windows Forms interface, no complex command-line arguments needed.
* **Timeline Analysis:** Filter files by `LastWriteTime`, `CreationTime`, or `LastAccessTime` within a specific date range.
* **MOTW Inspection:** Deep parsing of NTFS Alternate Data Streams (`Zone.Identifier`) to reveal:
    * **ZoneId** (e.g., Internet, Trusted).
    * **HostUrl** (Where the file was downloaded from).
    * **ReferrerUrl** ( The referring page).
* **Ownership Analysis:** Scan and filter files based on NTFS Owner (ACL), useful for identifying files created by specific compromised accounts.
* **Exclusion Lists:** Load a `.txt` file to exclude specific directories or system paths to speed up scanning.
* **Responsiveness:** Stop button allows for immediate cancellation of long-running scans (Fixed in v8.0).
* **Export:** Export results to CSV for further analysis in Excel or other tools.

## 🆕 What's New in v8.0
* **[Critical] UI Fix:** The "Stop" button is now fully responsive. The scanning loop uses a counter-based UI refresh mechanism to prevent the interface from freezing during heavy operations.
* **Exclusion Logic:** Fixed a bug where the "Pre-scan Owners" function was ignoring the exclusion list.

## 📋 Requirements
* Windows OS (Windows 10/11/Server recommended).
* PowerShell 5.1 (Default on Windows) or higher.
* Administrator privileges are recommended to access system files and read specific ACLs.

## 🛠️ How to Use

1.  **Download:** Clone this repository or download the `.ps1` script.
2.  **Run:** Right-click the script and select "Run with PowerShell".
    * *Note: You may need to set execution policy if you haven't already:*
      ```powershell
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
        ```
3.  **Interface Guide:**
    * **1. Target:** Select the drive or folder you want to analyze.
    * **2. Exclusions (Optional):** Load a text file containing paths to ignore (e.g., `C:\Windows\`).
    * **3. Owner Analysis:** Click "Pre-scan Owners" to populate the dropdown, then select a specific user to filter by.
    * **4. Time & MOTW:** Set your date range and ensure "Deep Parse MOTW" is checked for internet artifact analysis.
    * **5. Search:** Click "Start Forensic Search".

## 📂 Exclusion List Format
Create a `.txt` file with one path per line. Lines starting with `#` are ignored.
Example:
```text
# Exclude system folders
C:\Windows\System32
C:\Program Files\Common Files
