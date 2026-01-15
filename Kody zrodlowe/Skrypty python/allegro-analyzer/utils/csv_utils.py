import csv
import pandas as pd


# ---
# Czyta i waliduje plik .csv
# Zwraca DataFrame, status walidacji i listę problemów
# przykład użycia:
# df, is_valid, issues = reader_csv (
#                          file_path        = "file.csv",       -- (wymagane) ścieżka do pliku
#                          sep              = ';',              -- (wymagane) separator - domyślnie średnik
#                          required_columns = ["Col1","Col2"],  -- (opcjonalne) lista nazw kolumny lub None 
#                          required_header  = True,             -- (opcjonalne) czy plik ma nagłówek
#                          case_sensitive   = True              -- (opcjonalne) czy nazwy kolumn są rozróżniane wielkością liter
#                        )    
# ---
def reader_csv(
    file_path,
    sep=';',
    required_columns=None,
    required_header=True,
    case_sensitive=True
):
    issues = []
    is_valid = True
    rows = []

    def normalize(col):
        return col if case_sensitive else col.strip().lower()

    with open(file_path, newline='', encoding='cp1250') as f:
        reader = csv.reader(f, delimiter=sep)

        try:
            first_row = next(reader)
        except StopIteration:
            return None, False, ["Plik CSV jest pusty"]

        # CSV Z NAGŁÓWKIEM
        if required_header:
            header = first_row
            expected_len = len(header)

            # Walidacja wymaganych kolumn
            if required_columns:
                header_map = {
                    normalize(h): h for h in header
                }
                required_norm = {
                    normalize(c) for c in required_columns
                }

                missing = required_norm - set(header_map.keys())
                if missing:
                    issues.append(
                        f"Brak wymaganych kolumn: {sorted(missing)}"
                    )
                    is_valid = False

            start_line = 2

        # CSV BEZ NAGŁÓWKA
        else:
            expected_len = len(first_row)
            header = [f"Column{i+1}" for i in range(expected_len)]
            rows.append(first_row)
            start_line = 2

        # WALIDACJA STRUKTURY
        for line_no, row in enumerate(reader, start=start_line):
            if len(row) != expected_len:
                issues.append(
                    f"Wiersz {line_no}: {len(row)} kolumn "
                    f"(oczekiwano {expected_len})"
                )
                is_valid = False
            rows.append(row)

    # BUDOWA DATAFRAME
    df = pd.DataFrame(rows, columns=header)

    return df, is_valid, issues
