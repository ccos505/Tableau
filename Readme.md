## Tableau Permissions Fetcher

This script is designed to connect to a Tableau Server, fetch permissions for various Tableau items (projects, workbooks, datasources, and flows), and output the data into an Excel file.

# Prerequisites

Python 3.7 or higher
Tableau Server with appropriate user permissions to access the desired data

# Installation

Clone this repository or download the script.

Install the required Python packages using pip:
`pip install openpyxl tableauserverclient python-dotenv pandas numpy argparse pytz`

# Setup

1.Create a .env file in the same directory as the script with the following variables:
TABLEAU_SERVER_URL=<your_tableau_server_url>
TABLEAU_USERNAME=<your_tableau_username>
TABLEAU_PASSWORD=<your_tableau_password>
TABLEAU_SITE_ID=<your_tableau_site_id> # Optional, leave empty if not applicable

2.Ensure you have the necessary permissions on the Tableau Server to fetch user and item information.

# Usage

Run the script from the command line with optional arguments for site ID and output file name.

# Command-Line Arguments

--siteId or -s : (Optional) Specify the site ID if you want to fetch data from a specific Tableau site.
--output or -o : (Optional) Specify the output file name for the Excel file. Default is output.xlsx.

Example
To run the script with default settings:
`python main.py`

To specify a site ID and output file name:
`python main.py --siteId your_site_id --output custom_output.xlsx`

# Output

The script will create an Excel file with a sheet named "Permissions". The sheet will contain a pivot table of the fetched permissions, formatted for clarity.

# Troubleshooting

1.Ensure that the .env file is correctly set up with your Tableau Server credentials.
2.Verify that the Tableau Server URL, username, password, and site ID (if used) are correct.
3.Ensure you have the necessary permissions on the Tableau Server to access the desired data.
