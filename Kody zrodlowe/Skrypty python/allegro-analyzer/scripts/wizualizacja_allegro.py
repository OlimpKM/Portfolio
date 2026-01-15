import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import sys
import os

# Dodaje katalog główny projektu do ścieżki Pythona
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from utils.csv_utils import reader_csv
from config.csv_products import col_NAME, col_PRICE, col_BRAND, col_REVIEWS, col_CATEGORY_id, csv_SEP

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PRODUCTS = os.path.join(BASE_DIR, "..", "processed", "products.csv")
PNG_PRICE_DISTRIBUTION = os.path.join(BASE_DIR, "..", "raw", "price_distribution.png")
PNG_TOP_BRANDS = os.path.join(BASE_DIR, "..", "raw", "top_brands.png")


df, is_valid, issues = reader_csv(CSV_PRODUCTS, sep=",", required_columns=[col_NAME, col_PRICE, col_BRAND, col_REVIEWS, col_CATEGORY_id]) 

if is_valid:
    print("Plik CSV jest poprawny.")
else:   
    print("Plik CSV zawiera błędy:")
    for issue in issues:
        print("-", issue)   
    exit(1)        

plt.figure(figsize=(10,5))
sns.histplot(df[col_PRICE], bins=20)
plt.title("Rozkład cen produktów")
plt.xlabel("Cena (PLN)")
plt.ylabel("Liczba produktów")
plt.savefig(PNG_PRICE_DISTRIBUTION)

plt.figure(figsize=(10,5))
top_brands = df[col_BRAND].value_counts().head(5)
sns.barplot(x=top_brands.index, y=top_brands.values)
plt.title("Top 5 marek")
plt.ylabel("Liczba produktow")
plt.savefig(PNG_TOP_BRANDS)
