# WSUS_HTML_REPORTING

# WSUS Reporting Tool
The WSUS Reporting Tool is a PowerShell script that automates the generation of a WSUS report in HTML format. The tool will gather essential information, such as the FQDN, WSUS port, and SSL configuration, to create a comprehensive report. Additionally, it includes scheduled task functionality and email sending to streamline the reporting process.

Features
Automated WSUS report generation in HTML format.
Automatically detects FQDN, WSUS port, and SSL configuration.
Integrates PS2HTMLTable plugin for generating HTML tables.
Easy installation and setup.
Installation
To get started with the WSUS Reporting Tool, follow these steps:

Clone the repository to your local machine:
bash
Copy code
git clone https://github.com/your-username/WSUS-Reporting-Tool.git
Navigate to the cloned repository and locate the "WSUS_SCRIPT" folder.

Place the "WSUS_SCRIPT" folder in a suitable location, such as "My Documents."

Open a PowerShell terminal and navigate to the "WSUS_SCRIPT" folder.

Run the "Install-modules.ps1" script to install required modules:

mathematica
Copy code
.\Install-modules.ps1
Once the modules are installed, run the "Install-reporting.ps1" script to set up the reporting:
mathematica
Copy code
.\Install-reporting.ps1
The script will prompt you for any necessary configurations during the installation process.

Usage
Once the WSUS Reporting Tool is installed and configured, you can easily generate a WSUS report in HTML format by launching the scheduled task named "SRV_Reporting." The scheduled task will automatically run the script, gather the WSUS information, and create the report in HTML.

You can also run the script manually by executing the "WSUS-Report.ps1" file in the "WSUS_SCRIPT" folder:

Copy code
.\WSUS-Report.ps1
Customization
The WSUS Reporting Tool is designed to be easily customizable to fit your specific needs. You can modify the "WSUS-Report.ps1" script to include additional information or customize the HTML report layout.

Troubleshooting
If you encounter any issues during the installation or usage of the WSUS Reporting Tool, please check the following:

Ensure that you have the required permissions to run scheduled tasks and execute PowerShell scripts.

Double-check the configurations provided during the installation process.

Review the script logs and error messages to identify potential issues.

If problems persist, feel free to open an issue on the GitHub repository for assistance.

License
This project is licensed under the MIT License. Feel free to use, modify, and distribute it as per the license terms.

Acknowledgments
This project relies on the PS2HTMLTable plugin for generating HTML tables. Special thanks to its developers for their contribution.
