# SocieteGenerale-Hackathon

This project automates the process of fetching and updating historical stock data for a specified symbol from the Alpha Vantage API. The data is then updated into an Excel file for further analysis.

## Setup and Installation

### Prerequisites

Ensure you have the following installed:
- Python 3.6+
- pip (Python package installer)

### Required Libraries

Install the required Python libraries by running:
```sh
pip install requests pandas openpyxl
```


### Clone the Repository

Clone this repository to your local machine using:

```sh
git clone https://github.com/MohanB07/SocieteGenerale-Hackathon.git
cd SocieteGenerale-Hackathon
```

# Project Configuration

## API Key

Obtain an API key from [Alpha Vantage](https://www.alphavantage.co/).

## Symbol

Set the stock symbol you wish to fetch data for.

Modify the following variables in the script:

```python
API_KEY = 'YOUR_ALPHA_VANTAGE_API_KEY'
SYMBOL = 'YOUR_STOCK_SYMBOL'
```

# Project Setup and Configuration

## Excel Setup

Ensure you have an Excel file named `FinancialData.xlsx` in the root directory of the project. This file will be used to store the fetched stock data.

## Script Explanation

### Fetching Stock Data

The function `fetch_stock_data` fetches historical stock data from the Alpha Vantage API and returns it as a pandas DataFrame.

### Updating Excel Sheet

The function `update_excel_with_stock_data` updates the specified Excel sheet with the fetched stock data.

### Preloading Past Data

The function `preload_past_two_months_data` preloads stock data for the past two months for initial setup or testing purposes.

### Main Function

The main function orchestrates the process, preloading past data initially and scheduling daily updates.

## How to Run

### Preload Data

To preload the past two months of stock data, run the script:

```sh
python finance_data.py
```
### Daily Updates
The script is designed to run continuously and update the stock data daily at midnight.

## Contributing
Feel free to contribute to this project by submitting issues or pull requests. For major changes, please open an issue first to discuss what you would like to change.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgements
* Alpha Vantage for providing the stock data API.
* Pandas for data manipulation.
* OpenPyXL for Excel file handling.


