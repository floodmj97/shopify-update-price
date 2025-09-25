"""
Shopify Product Price Updater with Error Logging
------------------------------------------------
This script reads an Excel file containing SKUs and their updated prices,
then connects to a Shopify store via the Admin API and updates the prices.

It generates a CSV log file with the result of each attempted update.

Requirements:
    pip install pandas openpyxl shopifyapi
"""

import pandas as pd
import time
import shopify
from datetime import datetime

# -----------------------------
# CONFIGURATION
# -----------------------------
SHOP_URL = "your-shop-name.myshopify.com"
API_VERSION = "2023-10"  # Example API version
ACCESS_TOKEN = "your-admin-api-access-token"

EXCEL_FILE = "price_updates.xlsx"     # Your Excel file
SKU_COLUMN = "SKU"                    # Column name in Excel for SKU
PRICE_COLUMN = "New Price"            # Column name in Excel for price
LOG_FILE = f"update_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"  # Timestamped log file


# -----------------------------
# STEP 1: CONNECT TO SHOPIFY
# -----------------------------
def connect_to_shopify():
    """
    Initializes a connection to Shopify using the Admin API.
    """
    session = shopify.Session(f"https://{SHOP_URL}/admin/api/{API_VERSION}", ACCESS_TOKEN)
    shopify.ShopifyResource.activate_session(session)


# -----------------------------
# STEP 2: READ EXCEL FILE
# -----------------------------
def read_price_updates(filename):
    """
    Reads an Excel file and returns a dictionary mapping SKU → New Price.
    """
    df = pd.read_excel(filename)
    df = df.dropna(subset=[SKU_COLUMN, PRICE_COLUMN])  # Drop incomplete rows
    return dict(zip(df[SKU_COLUMN].astype(str), df[PRICE_COLUMN]))


# -----------------------------
# STEP 3: FIND PRODUCT VARIANT BY SKU
# -----------------------------
def find_variant_by_sku(sku):
    """
    Searches Shopify for a product variant matching the given SKU.
    Returns the variant object if found, otherwise None.
    """
    variants = shopify.Variant.find(sku=sku)  # Direct search by SKU
    if variants:
        return variants[0]
    return None


# -----------------------------
# STEP 4: UPDATE PRICE
# -----------------------------
def update_variant_price(variant, new_price, retries=3, delay=2):
    """
    Updates the price of a given variant in Shopify.
    Retries the update 'retries' times if it fails.
    
    Args:
        variant: The Shopify Variant object.
        new_price: The new price (float or string).
        retries: How many times to retry on failure.
        delay: Seconds to wait between retries.
    """
    for attempt in range(1, retries + 1):
        variant.price = str(new_price)  # Shopify expects price as a string
        success = variant.save()
        if success:
            return True
        else:
            print(f"⚠️ Attempt {attempt} failed for SKU {variant.sku}. Retrying...")
            time.sleep(delay)
    return False  # If all retries fail


# -----------------------------
# STEP 5: MAIN WORKFLOW
# -----------------------------
def main():
    # Connect to Shopify
    connect_to_shopify()
    print("Connected to Shopify API.")

    # Read price updates
    updates = read_price_updates(EXCEL_FILE)
    print(f"Loaded {len(updates)} price updates from Excel.")

    # Prepare log list
    log_entries = []

    # Loop through SKUs and update
    for sku, new_price in updates.items():
        print(f"Processing SKU: {sku} → {new_price}")
        variant = find_variant_by_sku(sku)

        if variant:
            success = update_variant_price(variant, new_price)
            if success:
                print(f"✅ Updated SKU {sku} to price {new_price}")
                log_entries.append({"SKU": sku, "Price": new_price, "Status": "Updated"})
            else:
                print(f"❌ Failed to update SKU {sku}")
                log_entries.append({"SKU": sku, "Price": new_price, "Status": "Failed to update"})
        else:
            print(f"⚠️ SKU {sku} not found in Shopify.")
            log_entries.append({"SKU": sku, "Price": new_price, "Status": "SKU not found"})

    # Save log to CSV
    log_df = pd.DataFrame(log_entries)
    log_df.to_csv(LOG_FILE, index=False)
    print(f"\n📄 Log file saved as: {LOG_FILE}")

    # End session
    shopify.ShopifyResource.clear_session()
    print("Finished updating prices.")


if __name__ == "__main__":
    main()
