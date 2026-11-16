<h1 align="center">Microsoft Office Installer</h1>

<p align="center">Open-source Windows tool to download and install Microsoft Office LTSC and Professional Plus editions using a single PowerShell command, featuring both a Graphical User Interface (GUI) and a Command Line Interface (CLI).</p>

<p align="center">
  <a href="https://learn.microsoft.com/powershell/"><img src="https://img.shields.io/badge/PowerShell-5.1-5391FE?style=flat-square" alt="PowerShell"></a>
  <a href="#"><img src="https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?style=flat-square" alt="Windows"></a>
  <a href="#"><img src="https://img.shields.io/badge/Zero-External%20Dependencies-2ea44f?style=flat-square" alt="Zero Dependencies"></a>
</p>

<hr>

## How to Install Microsoft Office / LTSC?

### Method 1 - PowerShell ❤️

1. Click the **Start Menu**, type `PowerShell`, right-click it, and select **Run as administrator**.

2. Copy and paste the command below and press **Enter**:

   - **GUI Edition (Recommended - Window interface):**
     ```powershell
     irm [https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-gui.ps1](https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-gui.ps1) | iex
     ```

   - **CLI Edition (Console interface):**
     ```powershell
     irm [https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-cli.ps1](https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-cli.ps1) | iex
     ```

3. Follow the interactive wizard on screen to choose your version, language, and preferred apps. Nothing installs without your final confirmation.

---

### Method 2 - Manual Execution

1. Download or clone this repository to your computer.
2. Open PowerShell as **Administrator** in the folder containing the scripts.
3. Run either `install-gui.ps1` or `install-cli.ps1`. *(Note: If script execution is restricted, run `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass` first).*

---

## Features

- **Dual Modes:** Graphical User Interface (`install-gui.ps1`) and Console (`install-cli.ps1`).
- **Supported Editions:** Office LTSC Professional Plus **2024** & **2021**; Office Professional Plus **2019**, **2016**, & **2013**.
- **Architecture:** Automatic detection for 64-bit or 32-bit systems.
- **Add-ons:** Optional inclusion for Project and Visio.
- **10 Languages:** English, Spanish, French, German, Brazilian Portuguese, Italian, Dutch, Polish, Russian, and Japanese.
- **Guided Workflow:** Every choice is explicitly selected on screen, and nothing is deployed without confirmation.
- **Modern GUI Design:** Features native system theme matching (Dark/Light mode) and branded app logo tiles.
- **Zero External Dependencies:** Automatically downloads the official Microsoft Office Deployment Tool (ODT) directly from Microsoft servers.

---

## Requirements

- **Windows 10 or 11**
- **PowerShell 5.0 or higher**
- **Administrator privileges**
- **.NET Framework 4.5 or higher**
- **Active Internet connection** (required to download ODT and Office setup files)

---

## Project Structure

```text
Microsoft-Office-Installer/
├── install-gui.ps1         # Graphical interface installer script
├── install-cli.ps1         # Command-line interface installer script
├── index.html              # Project presentation web page
└── README.md               # Documentation
