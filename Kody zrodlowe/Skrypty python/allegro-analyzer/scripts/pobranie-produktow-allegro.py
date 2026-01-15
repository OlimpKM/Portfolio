import asyncio
import aiohttp
import csv
import os
import pickle
import base64
import logging
from tqdm import tqdm
from collections import deque

# --- KONFIGURACJA ---
CLIENT_ID = os.getenv("ALLEGRO_CLIENT_ID")
CLIENT_SECRET = os.getenv("ALLEGRO_CLIENT_SECRET")
REDIRECT_URI = os.getenv("ALLEGRO_REDIRECT_URI")  # np. "https://twojlogin.github.io/nazwa-repozytorium/"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_LEAF = os.path.join(BASE_DIR, "..", "processed", "allegro_categories_leaf.csv")
CSV_PRODUCTS = os.path.join(BASE_DIR, "..", "processed", "allegro_products.csv")
CHECKPOINT_FILE = os.path.join(BASE_DIR, "..", "checkpoint", "products_checkpoint.pkl")
TOKEN_FILE = os.path.join(BASE_DIR, "..", "checkpoint", "allegro_token.pkl")

BATCH_SIZE = 10
REQUEST_DELAY = 0.2

# logowanie
LOG_FILE = os.path.join(BASE_DIR, "..", "log", "allegro_products.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# --- TOKEN ---
async def get_access_token(session, authorization_code=None, refresh_token=None):
    url = "https://allegro.pl.allegrosandbox.pl/auth/oauth/token"
    headers = {
        "Authorization": "Basic " + base64.b64encode(f"{CLIENT_ID}:{CLIENT_SECRET}".encode()).decode(),
        "Content-Type": "application/x-www-form-urlencoded"
    }

    if authorization_code:
        data = {
            "grant_type": "authorization_code",
            "code": authorization_code,
            "redirect_uri": REDIRECT_URI
        }
    elif refresh_token:
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token
        }
    else:
        raise RuntimeError("Nie podano authorization_code ani refresh_token")

    async with session.post(url, headers=headers, data=data) as resp:
        text = await resp.text()
        if resp.status != 200:
            raise RuntimeError(f"OAuth error {resp.status}: {text}")
        payload = await resp.json()
        # zapis tokena
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump({
                "access_token": payload["access_token"],
                "refresh_token": payload.get("refresh_token")
            }, f)
        return payload["access_token"], payload.get("refresh_token")

def load_token_file():
    """Wczytuje access + refresh token, jeśli plik istnieje"""
    if not os.path.exists(TOKEN_FILE):
        return None, None
    with open(TOKEN_FILE, "rb") as f:
        data = pickle.load(f)
    return data.get("access_token"), data.get("refresh_token")

# --- CSV i checkpoint ---
def load_categories():
    ids = []
    with open(CSV_LEAF, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            ids.append(row["id"])
    return ids

def append_products_batch(rows):
    print(f"Zapisuję {len(rows)} produktów do CSV")
    os.makedirs(os.path.dirname(CSV_PRODUCTS), exist_ok=True)
    file_exists = os.path.exists(CSV_PRODUCTS)
    with open(CSV_PRODUCTS, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["name","price","brand","reviews","category_id"])
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)

def save_checkpoint(queue):
    with open(CHECKPOINT_FILE, "wb") as f:
        pickle.dump(list(queue), f)

def load_checkpoint():
    if not os.path.exists(CHECKPOINT_FILE):
        return deque()
    with open(CHECKPOINT_FILE, "rb") as f:
        state = pickle.load(f)
    return deque(state)

# --- FETCH PRODUKTÓW ---
async def fetch_products(session, token, category_id, limit=100, offset=0):
    url = f"https://allegro.pl.allegrosandbox.pl/offers/listing?category.id={category_id}&limit={limit}&offset={offset}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.allegro.public.v1+json"
    }

    async with session.get(url, headers=headers) as resp:
        text = await resp.text()
        if resp.status == 401:
            # access_token wygasł
            raise RuntimeError("Token wygasł")
        elif resp.status in (403, 422):
            logging.warning(f"Pominięto kategorię {category_id} (status {resp.status})")
            return []
        elif resp.status != 200:
            logging.error(f"Błąd {resp.status}: {text}")
            return []

        data = await resp.json()
        items = []
        for offer in data.get("items", {}).get("regular", []):
            items.append({
                "name": offer.get("name"),
                "price": offer.get("sellingMode", {}).get("price", {}).get("amount"),
                "brand": offer.get("parameters", [{}])[0].get("values", [None])[0],
                "reviews": offer.get("reviews", {}).get("count"),
                "category_id": category_id
            })
        return items

# --- MAIN FETCH ---
async def fetch_all_products(authorization_code=None):
    queue = load_checkpoint()
    if not queue:
        categories = load_categories()
        queue = deque(categories)

    buffer = []

    async with aiohttp.ClientSession() as session:
        # Wczytaj istniejący token lub wygeneruj nowy
        access_token, refresh_token = load_token_file()
        if not access_token:
            if not authorization_code:
                raise RuntimeError("Nie ma tokena ani authorization_code")
            access_token, refresh_token = await get_access_token(session, authorization_code=authorization_code)

        with tqdm(total=len(queue), desc="Pobieram produkty") as pbar:
            while queue:
                category_id = queue.popleft()
                try:
                    items = await fetch_products(session, access_token, category_id)
                except RuntimeError as e:
                    if "Token wygasł" in str(e) and refresh_token:
                        # odśwież token tylko gdy wygasł
                        access_token, refresh_token = await get_access_token(session, refresh_token=refresh_token)
                        items = await fetch_products(session, access_token, category_id)
                    else:
                        raise

                if len(items) > 0:
                    print(f"Kategoria {category_id}: pobrano {len(items)} produktów")
                buffer.extend(items)
                pbar.update(1)

                if len(buffer) >= BATCH_SIZE:
                    append_products_batch(buffer)
                    save_checkpoint(queue)
                    logging.info(f"Zapisano batch {len(buffer)} produktów")
                    buffer.clear()

                await asyncio.sleep(REQUEST_DELAY)

            if buffer:
                append_products_batch(buffer)
                logging.info(f"Zapisano końcowy batch {len(buffer)} produktów")

# --- URUCHOMIENIE ---
if __name__ == "__main__":
    import sys
    authorization_code = sys.argv[1] if len(sys.argv) > 1 else None
    try:
        asyncio.run(fetch_all_products(authorization_code=authorization_code))
        print("✅ Pobieranie produktów zakończone")
    except KeyboardInterrupt:
        logging.warning("Proces przerwany przez użytkownika")
        print("⛔ Proces przerwany")
    except Exception as e:
        logging.exception("Błąd krytyczny")
        raise
