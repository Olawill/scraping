import requests
import json
import pandas as pd

# Defining the dictionart to hold the data of interest
all_products = {}

# Page range is determined from the api for the pages at 100 products per page
for page in range(1,228):
  url = f"https://www.bestbuy.ca/api/v2/json/sku-collections/16077?categoryid=&currentRegion=ON&include=facets%2C%20redirects&lang=en-CA&page={page}&pageSize=100&path=&query=&exp=labels%2Csearch_abtesting_5050_conversion%3Aa&sortBy=relevance&sortDir=desc"

  headers = {
    'authority': 'www.bestbuy.ca',
    'accept': '*/*',
    'accept-language': 'en-US,en;q=0.9',
    'cache-control': 'no-cache',
    'pragma': 'no-cache',
    'referer': 'https://www.bestbuy.ca/en-ca/collection/tv-home-theatre-on-sale/16077?icmp=global_shopbycategory_mth_ht',
    'sec-ch-ua': '"Google Chrome";v="107", "Chromium";v="107", "Not=A?Brand";v="24"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Mobile Safari/537.36',
    'x-dtpc': '-36$200223176_882h32vAJHNUFBQSKULFODCRWUDKFCHHWTJDAEV-0e0'
  }

  response = requests.request("GET", url, headers=headers)
  
  # Iterate throgh all the products on each page and add it to the all-products dictionary
  for product in response.json()['products']:

    all_products[product['sku']] = {
                                     'name': product['name'],
                                     'url': 'https://www.bestbuy.ca' + product['productUrl'],
                                     'customerRating': product['customerRating'],
                                     'customerRatingCount': product['customerRatingCount'],
                                     'customerReviewCount': product['customerReviewCount'],
                                     'regularPrice': product['regularPrice'],
                                     'salePrice': product['salePrice'],
                                     'categoryName': product['categoryName']
                                    }

# Convert the dictionary into a dataframe
tvs_accessories = pd.DataFrame.from_dict(all_products, orient='index')

# Convert the dataframe into an excel workbook
tvs_accessories.to_excel(r'C:\Users\OlawaleA\Downloads\archive\tvs_accessories.xlsx')
