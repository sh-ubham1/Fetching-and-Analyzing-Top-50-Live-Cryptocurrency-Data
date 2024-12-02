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

# Function to perform data analysis
def analyze_crypto_data(crypto_df):
    # Top 5 cryptocurrencies by market capitalization
    top_5_by_market_cap = crypto_df.nlargest(5, "Market Capitalization")

    # Average price of the top 50 cryptocurrencies
    average_price = crypto_df["Current Price (USD)"].mean()

    # Highest and lowest 24-hour percentage price change
    highest_price_change = crypto_df.nlargest(1, "Price Change (24h %)")[["Cryptocurrency Name", "Price Change (24h %)"]]
    lowest_price_change = crypto_df.nsmallest(1, "Price Change (24h %)")[["Cryptocurrency Name", "Price Change (24h %)"]]

    # Print analysis
    print("Top 5 Cryptocurrencies by Market Capitalization:")
    print(top_5_by_market_cap[["Cryptocurrency Name", "Market Capitalization"]])
    print("\nAverage Price of Top 50 Cryptocurrencies (USD):", average_price)
    print("\nHighest 24-hour Price Change:")
    print(highest_price_change)
    print("\nLowest 24-hour Price Change:")
    print(lowest_price_change)

# Main script
if __name__ == "__main__":
    crypto_df = fetch_crypto_data()
    if crypto_df is not None:
        print("Fetched Cryptocurrency Data:\n", crypto_df.head())
        analyze_crypto_data(crypto_df)
        # Save data to an Excel file
        crypto_df.to_excel("crypto_data_analysis.xlsx", index=False)
        print("Data saved to 'crypto_data_analysis.xlsx'.")
