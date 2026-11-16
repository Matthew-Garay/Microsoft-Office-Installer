<div align="center">

<img src="https://img.icons8.com/color/96/microsoft-office-2019.png" alt="Office Logo" class="logo-office" width="80" height="80">

# Microsoft Office Installer

 </div>

<p align="center">
  <b>A tool for Windows that allows downloading and deploying Microsoft Office (from LTSC 2024 down to 2013) through a single PowerShell command; it includes both graphical and console interfaces and requires no external dependencies.</b>
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

## System Requirements

<ul>
    <li><b>Operating System:</b> Windows 10 or Windows 11 (64-bit or 32-bit architecture).</li>
    <li><b>PowerShell:</b> Version 5.0 or higher.</li>
    <li><b>Permissions:</b> <b>Administrator</b> privileges (required for deploying Office packages).</li>
    <li><b>Internet Connection:</b> Required to securely download official setup files.</li>
</ul>

## Project Structure

<div style="background-color: #f6f8fa; border: 1px solid #d0d7de; border-radius: 6px; padding: 16px; font-family: ui-monospace, SFMono-Regular, SFMono, Consolas, 'Liberation Mono', Menlo, monospace; font-size: 85%; line-height: 1.45; color: #24292e;">
    <div style="font-weight: 600; margin-bottom: 8px;">Microsoft-Office-Installer/</div>
    <div style="padding-left: 16px;">
        ├── install-gui.ps1 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; # Main script with Graphical Interface (GUI)<br>
        ├── install-cli.ps1 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; # Interactive Command Line Interface (CLI) script<br>
        ├── index.html &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; # Project presentation web page<br>
        ├── LICENSE &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; # MIT License file<br>
        └── README.md &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; # Official repository documentation
    </div>
</div>

## Author & Contact

Developed by **Matthew Garay**.

<br>

<p>
  <a href="mailto:1911.garay.perez@gmail.com" title="Gmail">
    <img src="https://img.shields.io/badge/Gmail-EA4335?style=for-the-badge&logo=gmail&logoColor=white" alt="Gmail">
  </a>
  &nbsp;&nbsp;
  <a href="https://github.com/Matthew-Garay" target="_blank" title="GitHub">
    <img src="https://img.shields.io/badge/GitHub-181717?style=for-the-badge&logo=github&logoColor=white" alt="GitHub">
  </a>
  &nbsp;&nbsp;
  <a href="https://www.linkedin.com/in/matthewgarayperez/" target="_blank" title="LinkedIn">
    <img src="https://img.shields.io/badge/LinkedIn-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white" alt="LinkedIn">
  </a>
</p>



