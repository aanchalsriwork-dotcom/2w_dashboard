import os
from dotenv import load_dotenv
import requests
from datetime import date
from psycopg2.extras import execute_values
from db_connect import get_connection
from fetchers import get_countries
from pprint import pprint
import pandas as pd

ind_code = "EXCHANGE_RATE"
# Indicator code for exchange rates
# API key for the ExchangeRate API is stored in an environment variable.
# The load_dotenv() function loads environment variables from a .env file into the program's environment, allowing access to the API key without hardcoding it into the script.
load_dotenv()

API_KEY = os.environ.get("EXCHANGERATE_API_KEY")


# Function to fetch exchange rates from the ExchangeRate API.
# If the API call is successful, it returns a dictionary of currency codes and their corresponding exchange rates.
# If the API call fails, it raises a RuntimeError with the error message.
def fetch_rates():
    url = f"https://v6.exchangerate-api.com/v6/{API_KEY}/latest/USD"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    df = pd.DataFrame(data["conversion_rates"].items(), columns=["currency", "rate"])

    if data.get("result") != "success":
        raise RuntimeError(f"API call failed: {data}")
    return df.to_csv(rf"C:\Users\91931\2w_dashboard\exchange_rates.csv", index=False)
    # returns dictionary with currency code as key and ex rate as value
    # Eg: {"USD": 1, "INR": 83.1, "EUR": 0.92, ...}
    
pprint(fetch_rates())