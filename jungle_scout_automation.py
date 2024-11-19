import gspread
from oauth2client.service_account import ServiceAccountCredentials
from junglescout import Client
from junglescout.models.parameters.marketplace import Marketplace

# Jungle Scout API configuration
API_KEY_NAME = 'Saba_Automation'
API_KEY = '2gv0j3RMjNc-_Nh5E_ZXABbVmg7fWneDcp_DMqUamvI'
MARKETPLACE = Marketplace.US
client = Client(api_key_name=API_KEY_NAME, api_key=API_KEY, marketplace=MARKETPLACE)

# Google Sheets configuration
GOOGLE_SHEET_NAME = 'Automation'  # Replace with your Google Sheet name

# Embedded service account credentials
SERVICE_ACCOUNT_CREDS = {
    "type": "service_account",
    "project_id": "jungle-scout-automation",
    "private_key_id": "c9fe07ca9232292ada79b4c19b94f78a3dcce2e8",
    "private_key": """-----BEGIN PRIVATE KEY-----
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCowm6jIRAXt3gm
cc6p3FHKBkU5+GTA27xOOFS7M6A1aII89Q58XBGgTt0WSGC22f9aEx08dtq43ITT
VBg4YFwjqQdy8M1dQ2J8Aa62jIVIbpOXabjQztPDVjvOWMdVfJf9Z/ekPDTq2UCh
NQiz+r56uRTSZSZ/yUYN3FIH7ieE7Lib/VV84Ma0g5sS5YJER58SBcF06ZOtx2EY
ywgdpTt12MPYZdHg57POeequq5RbheSWw6hrnf+PrfrpmBVYGFmS8n58+eHSvUJN
5k8dxl7WyuEdckmQBUMcu8XsP+iItR/q+gYJgICngqJqEVYJCFygKJTThEv3VZj+
aQv2B6KnAgMBAAECggEAMZybTuzA4OAIe/DYKxqAlu45+r1ZzVKr7Kjly/T0486Z
/xah0BB3BBQ7LbpQaGD1D0NwRF7ClTcK+1+NEUHOxJMEBBfjn8fYE5GhDbyI0E7l
p80mToiHO7VFtrdlvm1290HRzSAABIyi0eGX8eVEiyqPAv1GkKmYpSmAmiE6z/oW
nhOonk3YgL6PUEA3vz2IKzIBgV+dS99heOY3KPNYUDM0khTMU92qfmLq1LGkI0I1
dnCCwu5a7IfWoF1JZszw+hEnO8eMFVc9CQoPjLCJl8MqLllDsPKhRZVAZn+7/TcH
T5uNdCMdXesV60x4W5iFPSWSF0yiNiWFUWER1aqVhQKBgQDokePozt2RUAukhN9h
+l9Ux6pSSL/0Zt/NBilo2FMB0BFqgt1Kvl/RLVkD8cgL8/E2hjF83cYyFIkbcJMr
t600GK87VVqOspxTHJ9cnfVgLIom1BlC+dqJQpO5uyMXE0B/nE/b82sJFPLcUjW/
RTU+f9GFITJL93VR2uTTMFuKewKBgQC5wtX2VjeTkKlPuMg2w8ZgI2+x20LdpdVd
EqfJY60Nwky3faf10OZ5DetpDTscmyGFwyjRqorCE0zaNVCZqZUNB2EAgaqv6z9d
KvpI8XYsGNb305b1FZnq/hJc6mOZkP/7shNW3Ksg8FSv8d06UwD0V5dcssF8rxYN
VxshK7aWxQKBgQCV8intWiLEo8U952VW+GQqdyk7MCiC3SkCOSzNqluqWYpBD+q5
XBDO/tvbjTGbc0ZcDx0tEpfMhiz4AhcBIsWLRzcDnD5srn/XniapQjaIMW9JJYq8
AfeCc+hm4V74a7M0E4Xxm/mwu07x+hcpIOf5SdO+b+7Lx9peUjEicJU0rQKBgET8
0d95Z4x7FhYQZvLHxF2h63JfHrcYRmQZcIy/Yt6QQVOH7B/DpERi4gGSs1hNWKbH
stoi/wNSjaEgWb2nmD5Ndj3s6goJUO/17Ru36Q45b2R8hTyh+BaoowM03SaEDj1Y
hgwlSbyi5KCvL1zgxKL6ALGhhXAbyhHMPrwT8uyNAoGBAI7BI9waRvHROLCoBZq7
cyUH7EcfMy1nOLDOe2EYaBFwKs/cI+n9RJXG63UjneQO4MXz+wxWdiZIfioamAij
jzsisylorXeBaN/q62G2XQKavwju/xs7oA6Xrdrv1dEm39tSaJRQd/KY78JuB6ge
l5ZGIw48RAFGDz5RlH/sBcRu
-----END PRIVATE KEY-----\n""",
    "client_email": "junglescout@jungle-scout-automation.iam.gserviceaccount.com",
    "client_id": "100766003402019318521",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/junglescout%40jungle-scout-automation.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

# Connect to Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(SERVICE_ACCOUNT_CREDS, scope)
client_gs = gspread.authorize(creds)


# Open the sheet
sheet = client_gs.open(GOOGLE_SHEET_NAME).sheet1

# Define headers for each data point
headers = [
    'Keyword', 'Exact Search Volume', 'Broad Search Volume', 'Ease of Ranking', 'Sponsored Product Count',
    'Average Monthly Sales', 'Average Amazon Revenue', 'Average Amazon Price', 'Average Reviews',
    'Highest Price Relevant Product', 'Highest Priced Product Units sold per month',
    'Highest Price #2 w/ More Than 100 Units/Month Sold', 'Highest #2 Units Sold per Month',
    'Highest Price #3 w/ More Than 100 Units/Month Sold', 'Highest #3 Units Sold per Month',
    '# of Sellers in Top 10 with less than 100 Reviews', '# of Irrelevant Listings on Page 1',
    'Monthly Trend', 'Quarterly Trend', 'Recommended Promotions', 'SP Brand Ad Bid', 'PPC Bid Broad', 'PPC Bid Exact',
    'Estimated 30-Day Search Volume'
]

# Add headers if they don't exist
if sheet.row_values(1) != headers:
    sheet.insert_row(headers, 1)

def safe_getattr(obj, attr, default=0):
    """Safely get an attribute from an object; if missing or None, return default."""
    value = getattr(obj, attr, default)
    return value if isinstance(value, (int, float, str)) else default

def is_row_empty(row_data):
    """Check if all fields in a row (except the keyword) are empty."""
    return all(not cell for cell in row_data[1:])  # Skip the 'Keyword' column

def fetch_keyword_insights(keyword, row_index):
    """Fetch data from the keyword insights endpoint and update Google Sheet."""
    try:
        response = client.keywords_by_keyword(search_terms=keyword, marketplace=MARKETPLACE)
        if not response.data:
            print(f"No keyword insight data returned for: {keyword}")
            return

        keyword_data = response.data[0].attributes
        sheet.update(f'B{row_index}', [[
            safe_getattr(keyword_data, 'monthly_search_volume_exact'),
            safe_getattr(keyword_data, 'monthly_search_volume_broad'),
            safe_getattr(keyword_data, 'ease_of_ranking_score'),
            safe_getattr(keyword_data, 'sponsored_product_count'),
            None, None, None, None,  # Placeholder for product data
            None, None, None, None,
            None, None, None, None,
            safe_getattr(keyword_data, 'monthly_trend'),
            safe_getattr(keyword_data, 'quarterly_trend'),
            safe_getattr(keyword_data, 'recommended_promotions'),
            safe_getattr(keyword_data, 'sp_brand_ad_bid'),
            safe_getattr(keyword_data, 'ppc_bid_broad'),
            safe_getattr(keyword_data, 'ppc_bid_exact'),
            safe_getattr(keyword_data, 'estimated_30_day_search_volume')
        ]])
        print(f"Keyword insight data for '{keyword}' updated in Google Sheet.")

    except Exception as e:
        print(f"Error fetching keyword insights for {keyword}: {e}")

def fetch_product_data(keyword, row_index):
    """Fetch data from the product database and update Google Sheet."""
    try:
        response = client.product_database(include_keywords=[keyword], marketplace=MARKETPLACE)
        if not response.data:
            print(f"No product database data returned for keyword: {keyword}")
            return

        total_sales, total_revenue, total_price, total_reviews = 0, 0, 0, 0
        highest_price, highest_price_units = 0, 0
        second_highest_price, second_highest_units = 0, 0
        third_highest_price, third_highest_units = 0, 0
        sellers_with_low_reviews = 0
        irrelevant_listings = 0
        count = 0

        for product in response.data:
            attributes = product.attributes
            units_sold = safe_getattr(attributes, 'approximate_30_day_units_sold', 0)
            revenue = safe_getattr(attributes, 'approximate_30_day_revenue', 0)
            price = safe_getattr(attributes, 'price', 0)
            reviews = safe_getattr(attributes, 'reviews', 0)

            total_sales += units_sold
            total_revenue += revenue
            total_price += price
            total_reviews += reviews
            count += 1

            if units_sold >= 100:
                if price > highest_price:
                    third_highest_price, third_highest_units = second_highest_price, second_highest_units
                    second_highest_price, second_highest_units = highest_price, highest_price_units
                    highest_price, highest_price_units = price, units_sold
                elif price > second_highest_price:
                    third_highest_price, third_highest_units = second_highest_price, second_highest_units
                    second_highest_price, second_highest_units = price, units_sold
                elif price > third_highest_price:
                    third_highest_price, third_highest_units = price, units_sold

            if reviews < 100:
                sellers_with_low_reviews += 1

        avg_sales = total_sales / count if count else 0
        avg_revenue = total_revenue / count if count else 0
        avg_price = total_price / count if count else 0
        avg_reviews = total_reviews / count if count else 0

        sheet.update(f'F{row_index}:Q{row_index}', [[
            avg_sales, avg_revenue, avg_price, avg_reviews,
            highest_price, highest_price_units,
            second_highest_price, second_highest_units,
            third_highest_price, third_highest_units,
            sellers_with_low_reviews, irrelevant_listings
        ]])
        print(f"Product data for '{keyword}' updated in Google Sheet.")

    except Exception as e:
        print(f"Error fetching product data for {keyword}: {e}")

def update_entire_sheet():
    """Update all rows in the Google Sheet."""
    rows = sheet.get_all_values()
    for idx, row in enumerate(rows[1:], start=2):  # Skip headers
        keyword = row[0]
        if keyword:
            print(f"Fetching data for keyword: {keyword}")
            fetch_keyword_insights(keyword, idx)
            fetch_product_data(keyword, idx)

def update_new_rows():
    """Update only rows with no data in the Google Sheet."""
    rows = sheet.get_all_values()
    for idx, row in enumerate(rows[1:], start=2):  # Skip headers
        keyword = row[0]
        if keyword and is_row_empty(row):
            print(f"Fetching data for new keyword: {keyword}")
            fetch_keyword_insights(keyword, idx)
            fetch_product_data(keyword, idx)

print("Select update mode:")
print("1: Update the entire sheet")
print("2: Update only new rows (rows with no data)")
choice = input("Enter your choice (1 or 2): ").strip()

if choice == '1':
    print("Updating the entire sheet...")
    update_entire_sheet()
elif choice == '2':
    print("Updating only new rows...")
    update_new_rows()
else:
    print("Invalid choice. Exiting.")
