import os
from dotenv import load_dotenv
import requests
from datetime import date
from psycopg2.extras import execute_values
from db_connect import get_connection
from fetchers import get_countries


ind_code = "004"
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
    if data.get("result") != "success":
        raise RuntimeError(f"API call failed: {data}")
    return data["conversion_rates"]
    # returns dictionary with currency code as key and ex rate as value
    # Eg: {"USD": 1, "INR": 83.1, "EUR": 0.92, ...}

def build_rows(countries: list[dict], rates: dict) -> list[tuple]:
    rows = []
    today = date.today().isoformat()
    
    for country in countries:
        
        # 1. Get the country's currency code
        currency_code = country["currency_code"]

        # 2. Find that currency's exchange rate
        rate = rates.get(currency_code)

        # 3. Skip if the API doesn't have this currency
        if rate is None:
            continue

        # 4. Create a database row using the WB country code
        row = (
        country["wb_code"],
        ind_code,
        today,
        rate
        )
        rows.append(row)
    return rows

def upsert_rates(conn, rows):
    sql = """
        INSERT INTO fact_macro (wb_code, ind_code, period_date, value)
        VALUES %s
        ON CONFLICT (wb_code, ind_code, period_date)
        DO UPDATE SET value = EXCLUDED.value, updated_at = now()
    """
    with conn.cursor() as cur:
        execute_values(cur, sql, rows)
    conn.commit()

def main():
    conn = get_connection()
    countries = get_countries(conn)
    rates = fetch_rates()
    rows = build_rows(countries, rates)
    upsert_rates(conn, rows)
    print(f"Upserted exchange rates for {len(rows)} countries")
    conn.close()
    
if __name__ == "__main__":
    main()
