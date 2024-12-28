import csv
import pandas as pd
from serpapi import GoogleSearch

# Configuration
SERPAPI_API_KEY = "EXAMPLE KEY"  # Replace with your SerpApi API key
STORE_IDS = [625, 1007, 1017, 635, 627]   # List of specific store IDs

def fetch_clearance_items(store_id, query="clearance", limit=10):
    """
    Fetch clearance items from a specific Home Depot store using SerpApi.
    """
    search_params = {
        "engine": "home_depot",
        "store_id": store_id,  # Specify the store ID
        "q": query,            # Search query (e.g., "clearance")
        "api_key": SERPAPI_API_KEY,
    }

    try:
        search = GoogleSearch(search_params)
        results = search.get_dict()
        products = results.get("products", [])[:limit]  # Limit the number of products
        return products
    except Exception as e:
        print(f"Error fetching items for store {store_id}: {e}")
        return []

def save_to_csv(items, filename="clearance_items.csv"):
    """
    Save items to a CSV file.
    """
    keys = ["store_id", "name", "price", "url"]
    with open(filename, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=keys)
        writer.writeheader()
        writer.writerows(items)
    print(f"Results saved to {filename}")

def save_to_excel(items, filename="clearance_items.xlsx"):
    """
    Save items to an Excel file using pandas.
    """
    df = pd.DataFrame(items)
    df.to_excel(filename, index=False, engine="openpyxl")
    print(f"Results saved to {filename}")

def display_items(items):
    """
    Display product details in a user-friendly format.
    """
    if not items:
        print("No products found.")
        return

    print(f"{'Store ID':<10} {'Name':<50} {'Price':<10} {'URL'}")
    print("-" * 120)
    for item in items:
        store_id = item.get("store_id", "Unknown")
        name = item.get("name", "Unknown")
        price = item.get("price", "Unknown")
        url = item.get("url", "Unknown")
        print(f"{store_id:<10} {name:<50} {price:<10} {url}")

if __name__ == "__main__":
    print(f"Searching for clearance items in stores: {STORE_IDS}\n")

    # Fetch clearance items for each store
    all_items = []
    for store_id in STORE_IDS:
        print(f"Fetching items for store ID: {store_id}...")
        items = fetch_clearance_items(store_id=store_id, query="clearance", limit=10)
        for item in items:
            # Format item data and add store_id
            all_items.append({
                "store_id": store_id,
                "name": item.get("title", "Unknown"),
                "price": item.get("price", "Unknown"),
                "url": item.get("link", "Unknown"),
            })

    # Display all results
    display_items(all_items)

    # Save results to CSV and Excel
    save_to_csv(all_items)
    save_to_excel(all_items)
