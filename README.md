# WCA World Directory Scraper

This Python script automates the process of scraping company data from the WCA World Directory website using Selenium and OpenPyXL. It extracts company names, names of contacts, emails, and phone numbers for a specified country.

## Prerequisites

Before running the script, ensure you have the following installed:
- Python (3.6 or higher)
- PyCharm IDE (optional but recommended for Python development)
- Google Chrome browser (recommended for Selenium)

## Installation

1. Clone this repository to your local machine:
   ```sh
   git clone [repository URL]
   ```
2. Install the required Python packages using pip:
   ```sh
   pip install -r requirements.txt
   ```

## Usage

1. Open the script in PyCharm or any text editor of your choice.
2. Update the following lines in the script with your WCA World Directory credentials:
   ```python
   username_input.send_keys("")  # add your username for the site
   password_input.send_keys("")  # add your own password
   ```
3. Customize the country and city selection as per your requirements:
   ```python
   country = "Australia"  # Adjust this value depending on the country data you want to scrape
   select.select_by_visible_text("All Cities")  # Update city selection if needed
   ```
4. Save the changes and run the script.
5. The script will open a Chrome browser, log in to the WCA World Directory website, and start scraping data.
6. The scraped data will be saved to an Excel file named `[country].xlsx` or appended to an existing Excel file `output1.xlsx`.

## Notes

- Ensure you have an active internet connection while running the script.
- The script uses the Chrome WebDriver. Make sure it is compatible with your Chrome browser version.
- Adjust the `time.sleep` intervals in the script as needed for smoother execution.
- For advanced usage, you can modify the script to scrape data for multiple countries or customize data extraction parameters.

Enjoy automating your data scraping tasks with this script!

