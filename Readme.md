# Tableau Permissions Fetcher <br>

This script is designed to connect to a Tableau Server, fetch permissions for various Tableau items (projects, workbooks, datasources, and flows), and output the data into an Excel file. <br>

# Prerequisites <br>

Python 3.7 or higher
Tableau Server with appropriate user permissions to access the desired data

# Installation <br>

Clone this repository or download the script. <br>

Install the required Python packages using pip: <br>
`pip install openpyxl tableauserverclient python-dotenv pandas numpy argparse pytz` <br>

# Setup <br>

1.Create a .env file in the same directory as the script with the following variables: <br>
`TABLEAU_SERVER_URL=<your_tableau_server_url>` <br>
`TABLEAU_USERNAME=<your_tableau_username>` <br>
`TABLEAU_PASSWORD=<your_tableau_password>` <br>
`TABLEAU_SITE_ID=<your_tableau_site_id> # Optional, leave empty if not applicable` <br>

2.Ensure you have the necessary permissions on the Tableau Server to fetch user and item information. <br>

# Usage <br>

Run the script from the command line with optional arguments for site ID and output file name. <br>

# Command-Line Arguments <br>

--siteId or -s : (Optional) Specify the site ID if you want to fetch data from a specific Tableau site. <br>
--output or -o : (Optional) Specify the output file name for the Excel file. Default is output.xlsx. <br>

Example <br>
To run the script with default settings: <br>
`python main.py` <br>

To specify a site ID and output file name: <br>
`python main.py --siteId your_site_id --output custom_output.xlsx` <br>

# Output <br>

The script will create an Excel file with a sheet named "Permissions". <br> The sheet will contain a pivot table of the fetched permissions, formatted for clarity. <br>

# Troubleshooting <br>

1.Ensure that the .env file is correctly set up with your Tableau Server credentials. <br>
2.Verify that the Tableau Server URL, username, password, and site ID (if used) are correct. <br>
3.Ensure you have the necessary permissions on the Tableau Server to access the desired data. <br>
