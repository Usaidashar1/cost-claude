"""
Diagnostic script - run this on a machine that can reach prices.azure.com
(this sandbox can't). It prints the raw Consumption + Reservation entries
for each flagged SKU/region so we can see exactly what the API returns,
rather than guessing.

Usage:
    python diagnose_sku.py "Standard_D2_v3" eastus INR
    python diagnose_sku.py "Standard_B2s" eastus INR
    python diagnose_sku.py "Standard_F2s" eastus INR
    python diagnose_sku.py "Standard_F2s_v2" eastus INR
    python diagnose_sku.py "Standard_F4s" eastus INR
    python diagnose_sku.py "Standard_F4s_v2" eastus INR
    python diagnose_sku.py "Standard_D4s_v3" eastus INR   # the "wrong pricing" one
"""
import sys
import json
import requests

API = "https://prices.azure.com/api/retail/prices"

def query(sku, region, currency, price_type):
    filt = f"armSkuName eq '{sku}' and armRegionName eq '{region}' and priceType eq '{price_type}'"
    r = requests.get(API, params={"api-version": "2023-01-01-preview", "$filter": filt, "currencyCode": currency}, timeout=15)
    r.raise_for_status()
    return r.json().get("Items", [])

def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)
    sku, region, currency = sys.argv[1], sys.argv[2], sys.argv[3]

    print(f"\n=== {sku} / {region} / {currency} ===\n")

    print("--- Consumption (PAYG) ---")
    payg = query(sku, region, currency, "Consumption")
    if not payg:
        print("  NO RESULTS - the armSkuName/armRegionName combo returned nothing at all.")
        print("  Try region names like 'eastus', 'centralindia', 'southindia', 'uksouth' (all lowercase, no spaces).")
    for i in payg:
        print(f"  productName={i.get('productName')!r}  meterName={i.get('meterName')!r}  "
              f"retailPrice={i.get('retailPrice')}  unitOfMeasure={i.get('unitOfMeasure')}  "
              f"armSkuName={i.get('armSkuName')!r}")

    print("\n--- Reservation (RI) ---")
    ri = query(sku, region, currency, "Reservation")
    if not ri:
        print("  NO RESULTS - this SKU/region has no Reservation entries under this exact armSkuName.")
        print("  If the Calculator DOES show an RI price, the SKU is likely filed under a slightly")
        print("  different armSkuName (check the armSkuName field above from the Consumption query -")
        print("  it should be identical for Reservation entries too).")
    for i in ri:
        print(f"  productName={i.get('productName')!r}  meterName={i.get('meterName')!r}  "
              f"reservationTerm={i.get('reservationTerm')}  retailPrice={i.get('retailPrice')}  "
              f"armSkuName={i.get('armSkuName')!r}")

    print()

if __name__ == "__main__":
    main()
