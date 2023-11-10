# WSUS Reporting Tool (+French Translation)

The WSUS Reporting Tool is a PowerShell script that automates the generation of a WSUS report in HTML format. The tool will gather essential information, such as the FQDN, WSUS port, and SSL configuration, to create a comprehensive report. Additionally, it includes scheduled task functionality and email sending to streamline the reporting process.

## Features

- Automated WSUS report generation in HTML format.
- Automatically detects FQDN, WSUS port, and SSL configuration.
- Integrates PS2HTMLTable plugin for generating HTML tables.
- Easy installation and setup.

## Installation

To get started with the WSUS Reporting Tool, follow these steps:

0. Log into the WSUS Server with an admin acc. (the script will add the necessary groups)

1. Clone the repository to your local machine:

``git clone https://github.com/GiZcesi/WSUS_HTML_REPORTING``

2. Navigate to the cloned repository and locate the "WSUS_SCRIPT" folder.

3. Place the "WSUS_SCRIPT" folder in "My Documents." on the WSUS Server.

4. Open a PowerShell terminal and navigate to the "WSUS_SCRIPT" folder.

5. Run "Install-modules.ps1" and "Install-Reporting.ps1"
The script will prompt you for any necessary configurations during the installation process.

## Usage

Once the WSUS Reporting Tool is installed and configured, you can easily generate a WSUS report in HTML format by launching the scheduled task named "SRV_Reporting." The scheduled task will automatically run the script, gather the WSUS information, and create the report in the "WSUS_SCRIPT\export" folder.


