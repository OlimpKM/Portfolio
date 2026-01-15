import asyncio
import aiohttp
import csv
import os
import base64
import logging
import pickle
from collections import deque
from tqdm import tqdm
from collections import deque

# konfiguracja
CLIENT_ID = "7a11357893584d11878615ea814f4fdc" #os.getenv("ALLEGRO_CLIENT_ID")
CLIENT_SECRET = "2Hvw7s2FxBUuqHjVLcJukInMEJFDiQv3FjMH7k74fdtYdZhIJyoLEq1bpTB07U4P" #os.getenv("ALLEGRO_CLIENT_SECRET")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_ALL = os.path.join(BASE_DIR,"..","processed", "allegro_categories_all.csv")
CSV_LEAF = os.path.join(BASE_DIR, "..","processed", "allegro_categories_leaf.csv")
LOG_FILE = os.path.join(BASE_DIR, "..","log","allegro_categories.log")
CHECKPOINT_FILE = os.path.join(BASE_DIR,"..","checkpoint", "categories_checkpoint.pkl")

BATCH_SIZE = 10           # ile rekordów zapisywać jednorazowo
REQUEST_DELAY = 0.2       # delikatny throttling

# logowanie
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# pobranie tokenu OAuth2
async def get_access_token(session: aiohttp.ClientSession) -> str:
    url = "https://allegro.pl/auth/oauth/token"

    auth = f"{CLIENT_ID}:{CLIENT_SECRET}"
    auth_b64 = base64.b64encode(auth.encode()).decode()

    headers = {
        "Authorization": f"Basic {auth_b64}",
        "Content-Type": "application/x-www-form-urlencoded"
    }

    data = {"grant_type": "client_credentials"}

    async with session.post(url, headers=headers, data=data) as resp:
        if resp.status != 200:
            text = await resp.text()
            logging.error(f"OAuth error {resp.status}: {text}")
            raise RuntimeError("Nie udało się pobrać tokenu")
        payload = await resp.json()
        return payload["access_token"]

# --- Sekcja CSV ---   
def load_existing_ids():
    ids = set()
    if not os.path.exists(CSV_ALL):
        return ids

    with open(CSV_ALL, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            ids.add(row["id"])
    return ids


def append_batch(rows):
    file_exists = os.path.exists(CSV_ALL)
    with open(CSV_ALL, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["id", "name", "parent_id"]
        )
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)

# --- Sekcja checkpointów ---
# zapis checkpointu
def save_checkpoint(queue, visited):
    state = {
        "queue": list(queue),
        "visited": list(visited)
    }
    with open(CHECKPOINT_FILE, "wb") as f:
        pickle.dump(state, f)

# wczytanie checkpointu
def load_checkpoint():
    if not os.path.exists(CHECKPOINT_FILE):
        return deque([None]), set()

    with open(CHECKPOINT_FILE, "rb") as f:
        state = pickle.load(f)

    return deque(state["queue"]), set(state["visited"])


# --- Sekcja kategorie ---
# pobranie kategorii (asynchronicznie)
async def fetch_categories(session, token, parent_id=None):
    url = "https://api.allegro.pl/sale/categories"
    if parent_id:
        url += f"?parent.id={parent_id}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.allegro.public.v1+json"
    }

    async with session.get(url, headers=headers) as resp:
        if resp.status != 200:
            text = await resp.text()
            logging.error(f"Categories error {resp.status}: {text}")
            return []
        data = await resp.json()
        return data.get("categories", [])

# pobranie wszystkich kategorii (asynchronicznie)
async def fetch_all_categories():
    queue, visited = load_checkpoint()    
    buffer = []

    async with aiohttp.ClientSession() as session:
        token = await get_access_token(session)

        with tqdm(
            desc="Pobieram kategorie",
            unit=" kat",
            initial=len(visited)
        ) as pbar:

            while queue:
                parent_id = queue.popleft()
                cats = await fetch_categories(session, token, parent_id)

                for cat in cats:
                    cid = cat.get("id")
                    if cid in visited:
                        continue

                    visited.add(cid)

                    parent = cat.get("parent")
                    parent_id = parent["id"] if isinstance(parent, dict) else None

                    record = {
                        "id": cid,
                        "name": cat.get("name"),
                        "parent_id": parent_id
                    }

                    buffer.append(record)
                    queue.append(cid)

                    pbar.update(1)  # nowa kategoria (progress bar)

                    if len(buffer) >= BATCH_SIZE:
                        append_batch(buffer)
                        save_checkpoint(queue, visited)
                        logging.info(f"Zapisano batch {len(buffer)}")
                        buffer.clear()

                    await asyncio.sleep(REQUEST_DELAY)

            if buffer:
                append_batch(buffer)
                logging.info(f"Zapisano końcowy batch {len(buffer)}")

# znajdz ostatnie kategorie (leaf) i zapisz do osobnego CSV
def build_leaf_csv():
    rows = []
    parents = set()

    with open(CSV_ALL, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
            if row["parent_id"]:
                parents.add(row["parent_id"])

    leaf_rows = [r for r in rows if r["id"] not in parents]

    with open(CSV_LEAF, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["id", "name", "parent_id"])
        writer.writeheader()
        writer.writerows(leaf_rows)

    logging.info(f"Utworzono CSV leaf: {len(leaf_rows)} rekordów")

# uruchomienie
if __name__ == "__main__":
    try:
        #asyncio.run(fetch_all_categories())
        build_leaf_csv()
        print("✅ Pobieranie kategorii zakończone")
    except KeyboardInterrupt:
        logging.warning("Proces przerwany przez użytkownika")
        print("⛔ Proces przerwany")
    except Exception as e:
        logging.exception("Błąd krytyczny")
        raise
