import requests
import pandas as pd

# Function to fetch live cryptocurrency data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        data = response.json()
        # Extracting required fields
        crypto_data = []
        for coin in data:
            crypto_data.append({
                "Cryptocurrency Name": coin["name"],
                "Symbol": coin["symbol"].upper(),
                "Current Price (USD)": coin["current_price"],
                "Market Capitalization": coin["market_cap"],
                "24h Trading Volume": coin["total_volume"],
                "Price Change (24h %)": coin["price_change_percentage_24h"]
            })
        return pd.DataFrame(crypto_data)
    else:
        print("Failed to fetch data. HTTP Status code:", response.status_code)
        return None

# Fetch and display data
if __name__ == "__main__":
    crypto_df = fetch_crypto_data()
    if crypto_df is not None:
        print(crypto_df)
        # Save data to an Excel file
        crypto_df.to_excel("crypto_data.xlsx", index=False)
        print("Data saved to 'crypto_data.xlsx'.")
