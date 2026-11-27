import os
from dotenv import load_dotenv
import psycopg2

# Load the variables stored in .env file
load_dotenv()


def get_connection():
    # Get Supabase database URL
    db_url = os.environ["SUPABASE_DB_URL"]

    # Connect to the database and return the connection object.
    return psycopg2.connect(db_url)

if __name__ == "__main__":
    conn = get_connection()
    print("Connection successful:", conn)
    conn.close()
