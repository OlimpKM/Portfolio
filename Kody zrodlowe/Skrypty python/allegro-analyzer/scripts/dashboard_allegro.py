import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
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

st.title("Allegro Product Trends Dashboard")
st.dataframe(df)

st.subheader("Rozkład cen")
fig, ax = plt.subplots()
sns.histplot(df[col_PRICE], bins=20, ax=ax)
st.pyplot(fig)

st.subheader("Top marki")
fig, ax = plt.subplots()
top_brands = df[col_BRAND].value_counts().head(5)
sns.barplot(x=top_brands.index, y=top_brands.values, ax=ax)
st.pyplot(fig)
