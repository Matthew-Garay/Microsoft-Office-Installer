<h1 align="center">Microsoft Office Installer</h1>

<p align="center">Open-source PowerShell deployment tool designed to automate the download and installation of Microsoft Office using the official Office Deployment Tool (ODT), featuring GUI and CLI modes with zero external dependencies.</p>

<hr>

## How to Run Microsoft Office Installer?

### Method 1 - PowerShell

1. Click the **Start Menu**, type `PowerShell`, and open it as **Administrator**.

2. Copy and paste one of the commands below and press **Enter**:

   - **Graphical Interface (GUI):**
     ```powershell
     irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/instalar-gui.ps1 | iex
     ```

   - **Command Line (CLI):**
     ```powershell
     irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/instalar-cli.ps1 | iex
     ```

3. In the menu that appears, select your preferred Office edition, architecture, language, and specific applications.

---

### Method 2 - Manual Download

1. Download the script directly from the repository:
   * [**instalar-gui.ps1**](https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/instalar-gui.ps1) (Graphical Interface)
   * [**instalar-cli.ps1**](https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/instalar-cli.ps1) (Command Line Interface)

2. Right-click the downloaded `.ps1` file and select **Run with PowerShell** (or execute it inside an elevated PowerShell session).

---

> [!TIP]
> - If script execution is restricted on your system, allow script execution for the current session by running `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` before running the command.
> - **Having trouble?** Verify system requirements or open an issue on the [GitHub repository](https://github.com/Matthew-Garay/Microsoft-Office-Installer/issues).

> [!NOTE]
> - The `irm` (Invoke-RestMethod) command downloads the script directly from the raw GitHub source, and `iex` (Invoke-Expression) executes it in memory.
> - Always double-check the URL source before executing remote scripts in PowerShell.
> - This tool automates official Microsoft Office Deployment Tool (ODT) setup directly from Microsoft servers. It does not include, distribute, or automate software licensing or activation. A valid Microsoft Office license is required.

---

## Key Features

* **Native Execution:** Built entirely using Windows PowerShell and WinForms without third-party runtimes.
* **Official Deployment Engine:** Downloads and executes the official Microsoft Office Deployment Tool (ODT) directly from Microsoft servers.
* **Granular Application Selection:** Individually include or exclude Word, Excel, PowerPoint, Outlook, Access, Publisher, OneNote, Skype for Business, Project, and Visio.
* **Multi-Version Support:** Deployment support for Office LTSC 2024, 2021, 2019, 2016, and 2013 across x86 and x64 architectures.
* **Dynamic GUI Theme:** The graphical interface automatically queries Windows registry keys to match Light or Dark system theme settings.

---

## System Requirements

| Component | Requirement |
| :--- | :--- |
| **Operating System** | Windows 10 / Windows 11 (32-bit or 64-bit) |
| **Shell Environment** | Windows PowerShell 5.0 or higher |
| **Framework** | .NET Framework 4.5 or higher |
| **Privileges** | Administrative privileges (Run as Administrator) |
| **Network** | Active internet connection |

---

<div align="center">

# Web Interface - [https://matthew-garay.github.io/Microsoft-Office-Installer/](https://matthew-garay.github.io/Microsoft-Office-Installer/)

Latest Version: v1.0

</div>
