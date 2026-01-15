import pandas as pd
import sys
import os

# Dodaje katalog główny projektu do ścieżki Pythona
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from utils.csv_utils import reader_csv
from config.csv_products import col_NAME, col_PRICE, col_BRAND, col_REVIEWS, col_CATEGORY_id, csv_SEP

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PRODUCTS = os.path.join(BASE_DIR, "..", "processed", "products.csv")

df, is_valid, issues = reader_csv(CSV_PRODUCTS, sep=",", required_columns=[col_NAME, col_PRICE, col_BRAND, col_REVIEWS, col_CATEGORY_id]) 

if is_valid:
    print("Plik CSV jest poprawny.")
else:   
    print("Plik CSV zawiera błędy:")
    for issue in issues:
        print("-", issue)   
    exit(1)        

df[col_PRICE] = df[col_PRICE].str.replace(",", ".").astype(float)

avg_price = df[col_PRICE].mean()
top_brands = df[col_BRAND].value_counts().head(5)
most_reviewed = df.sort_values(col_REVIEWS, ascending=False).head(5)

print("Średnia cena:", round(avg_price,2))
print("\nTop 5 marek:" )
for brand, count in top_brands.items():  print(f"- {brand} ({count} produktów)")

print("\nNajwięcej opinii")
for _, row in most_reviewed.iterrows():
    print(f"{row[col_NAME]} - {row[col_REVIEWS]} opinii")