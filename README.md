AD Management Tool
A PowerShell-based GUI application for managing Active Directory users, distribution groups, and contacts with email integration.
Overview
The AD Management Tool provides a tabbed interface to simplify common Active Directory tasks:
User Creation: Create AD users with customizable naming patterns, email addresses, and additional attributes.

Distribution Group Management: Create or update universal distribution groups with email features.

Contact Creation: Add external contacts to AD with email and company details.

Features
Intuitive GUI with three tabs for Users, Groups, and Contacts.

Bulk user import from CSV files.

Email configuration for Exchange integration (optional).

Customizable settings via a configuration dialog.

Logging for all actions (stored in %USERPROFILE%\Documents\AD_Management_Logs).

Dark mode toggle and menu options for tools like ADUC and Exchange Admin Center.

Prerequisites
Operating System: Windows (with PowerShell 5.1 or later).

Modules: Active Directory PowerShell module (RSAT-AD-PowerShell).

Privileges: Administrative rights required for AD operations.

Assemblies: .NET Framework (for Windows Forms).

Installation
Clone or Download:
bash

git clone https://github.com/MoshikoKar/AD-Management-Tool.git

Or download the ZIP and extract it.

Install Required Module:
Ensure the Active Directory module is installed:
powershell

Install-WindowsFeature RSAT-AD-PowerShell

Run this command in an elevated PowerShell prompt if the module is missing.

Run the Script:
Open PowerShell as Administrator.

Navigate to the script directory:
powershell

cd path\to\AD-Management-Tool

Execute:
powershell

.\AD-Management-Tool.ps1

Usage
Launch the Tool:
Run the script with administrative privileges. If not elevated, it will prompt to restart as admin.

Configure Settings (optional):
Go to File > Configuration... to customize:
Organization details (company name, domain).

OU paths for users, groups, and contacts.

Email settings (Exchange server, username pattern).

Save changes to persist settings in %APPDATA%\AD-Management-Tool\config.xml.

Tabs:
Create User:
Fill in required fields (First Name, Last Name, Primary Domain).

Optional: Add additional domains, description, etc.

Click "Create User" or "Bulk Import" for CSV imports.

Manage Distribution Group:
Enter group details and domain.

Click "Create/Update Mail Group" to create or modify.

Create Contact:
Input contact info (First Name, Last Name, Email, etc.).

Click "Create Contact".

View Logs:
Click "View Log" on any tab or "Open Logs Folder" at the bottom to review actions.

Tools Menu:
Access ADUC, Exchange Admin Center, or generate a CSV template for bulk imports.

CSV Bulk Import Format
For bulk user imports, use this CSV structure:
csv

FirstName,LastName,PrimaryDomain,AdditionalDomains,Description,Office,Department,JobTitle
John,Doe,domain.com,other.com;third.com,IT Staff,Main Office,IT,Admin
Jane,Smith,domain.com,,HR Staff,Branch Office,HR,Manager

Required: FirstName, LastName, PrimaryDomain.

Optional: Other fields (use semicolons for multiple AdditionalDomains).

Configuration
Settings are stored in %APPDATA%\AD-Management-Tool\config.xml. Default values are auto-detected where possible (e.g., domain from environment variables). Modify via the GUI or edit the XML directly.
Logs
All actions are logged to:
%USERPROFILE%\Documents\AD_Management_Logs\AD_User_Creation.log

%USERPROFILE%\Documents\AD_Management_Logs\AD_Group_Management.log

%USERPROFILE%\Documents\AD_Management_Logs\AD_Contact_Creation.log

Troubleshooting
Module Missing: Install RSAT-AD-PowerShell if AD cmdlets fail.

Permission Errors: Ensure the script runs as Administrator.

GUI Issues: Verify .NET Framework is installed and PowerShell execution policy allows scripts (Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass).

Contributing
Fork the repository.

Create a feature branch (git checkout -b feature-name).

Commit changes (git commit -m "Add feature").

Push to the branch (git push origin feature-name).

Open a Pull Request.

License
This project is licensed under the MIT License - see the LICENSE file for details.
Author
Claude

GitHub: https://github.com/MoshikoKar

Version
1.0.0 (Initial Release)
