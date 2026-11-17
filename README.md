<div align="center">

<img src="https://img.icons8.com/color/96/microsoft-office-2019.png" alt="Office Logo" class="logo-office" width="80" height="80">

# Microsoft Office Installer

 </div>

<p align="center">
  <b>A tool for Windows that allows downloading and deploying Microsoft Office (from LTSC 2024 down to 2016) through a single PowerShell command; it includes both graphical and console interfaces and requires no external dependencies.</b>
</p>
 
## About the Project

**Microsoft Office Installer** is an automated and modern solution developed in PowerShell that completely simplifies the deployment of Microsoft productivity suites for home or professional environments on Windows systems.

<p align="center">
  <img src="https://img.icons8.com/color/48/microsoft-word-2019.png" alt="Word" title="Microsoft Word" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/microsoft-excel-2019.png" alt="Excel" title="Microsoft Excel" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/microsoft-powerpoint-2019.png" alt="PowerPoint" title="Microsoft PowerPoint" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/microsoft-outlook-2019.png" alt="Outlook" title="Microsoft Outlook" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/microsoft-access-2019.png" alt="Access" title="Microsoft Access" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/microsoft-publisher-2019.png" alt="Publisher" title="Microsoft Publisher" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/microsoft-onenote-2019.png" alt="OneNote" title="Microsoft OneNote" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/skype.png" alt="Skype" title="Skype for Business" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/microsoft-project-2019.png" alt="Project" title="Microsoft Project" width="40" height="40">
  &nbsp;&nbsp;
  <img src="https://img.icons8.com/color/48/microsoft-visio-2019.png" alt="Visio" title="Microsoft Visio" width="40" height="40">
</p>

---

## How to Install

### Method 1: Direct PowerShell Execution 

1. Open the **Start Menu**, type **PowerShell**, right-click the application, and select **Run as administrator**.
2. *(Optional - First time only)* If your system restricts script execution, enable it by running this command first:

   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass
Console Edition (CLI):
   ```powershell
   irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-cli.ps1 | iex
 ```
Graphical Edition (GUI - Recommended):
   ```powershell
  irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-gui.ps1 | iex
 ```

### Method 2: Local / Downloaded Execution

1. Clone or download this repository as a **.zip** file and extract it to your computer.
2. Open PowerShell as **Administrator** and navigate to the extracted folder directory where the scripts are located.
  Run the local script depending on your preference:

For the graphical interface:
 ```powershell
    .\install-gui.ps1
 ```
For the console interface:
```powershell
 .\install-cli.ps1
 ```

---

## System Requirements

<table align="center">
  <tr>
    <th align="left" width="220">Requirement</th>
    <th align="left">Details</th>
  </tr>
  <tr>
    <td><b>Operating System</b></td>
    <td>Windows 10 or Windows 11 (64-bit or 32-bit architecture)</td>
  </tr>
  <tr>
    <td><b>PowerShell</b></td>
    <td>Version 5.0 or higher (included with Windows)</td>
  </tr>
  <tr>
    <td><b>Permissions</b></td>
    <td>Administrator privileges (required for deploying Office packages)</td>
  </tr>
  <tr>
    <td><b>Internet Connection</b></td>
    <td>Required to securely download the official setup files from Microsoft</td>
  </tr>
</table>

---

## Project Structure

```text
Microsoft-Office-Installer/
│
├── install-gui.ps1    →  Graphical installer (GUI)
│                         Version, language and application selection
│
├── install-cli.ps1    →  Console installer (CLI)
│                         Interactive step-by-step installation
│
├── index.html         →  Project presentation web page
│                         Multilingual (English, Spanish, French, German)
│
├── LICENSE            →  MIT License
│
└── README.md          →  Official repository documentation
```

---

## Contact

<div align="center">

[![Gmail](https://skillicons.dev/icons?i=gmail)](mailto:1911.garay.perez@gmail.com "Gmail")&nbsp;&nbsp;&nbsp;[![GitHub](https://skillicons.dev/icons?i=github)](https://github.com/Matthew-Garay "GitHub")&nbsp;&nbsp;&nbsp;[![LinkedIn](https://skillicons.dev/icons?i=linkedin)](https://www.linkedin.com/in/matthewgarayperez/ "LinkedIn")

</div>

---

<div align="center">
  <sub><b>Microsoft Office Installer</b> &nbsp;·&nbsp; Version 1.0 &nbsp;·&nbsp; MIT License</sub>
</div>
