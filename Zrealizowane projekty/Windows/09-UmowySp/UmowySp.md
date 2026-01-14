# UmowySp
**Enterprise Sales Documents Archiving System**

![Status](https://img.shields.io/badge/Status-Zrealizowany-success)  
![Technology](https://img.shields.io/badge/Technologia-C%23%20%2F%20WinForms%20%2F%20SQL%20Server-blue)  
![Deployment](https://img.shields.io/badge/Deployment-ClickOnce-informational)  
![Auth](https://img.shields.io/badge/Autoryzacja-Active%20Directory-critical)  
![Legal](https://img.shields.io/badge/Typ-Projekt%20komercyjny%20(UoP)-orange)  
![Owner](https://img.shields.io/badge/Prawa%20autorskie-Distribev-lightgrey)  

**UmowySp** to wewnętrzny system klasy **back-office**, przeznaczony do **archiwizacji, ewidencji oraz obsługi dokumentów handlowych** (umów sprzedażowych i dokumentów powiązanych) w środowisku korporacyjnym.

Aplikacja została zaprojektowana jako **lekki, stabilny klient desktopowy**, zintegrowany z infrastrukturą domenową Windows oraz centralną bazą danych **Microsoft SQL Server**, zapewniając spójność danych, kontrolę dostępu i pełną audytowalność operacji.

---

## 🔐 Autoryzacja i bezpieczeństwo

- Logowanie użytkowników odbywa się poprzez **Active Directory (Windows Authentication)**
- Użytkownik loguje się **tymi samymi poświadczeniami, co do systemu operacyjnego**
- System identyfikuje użytkownika na podstawie **klucza SID domenowego**
- Przynależność do grup AD determinuje rolę w systemie:

---

## 🚀 Kluczowe funkcjonalności

### 1. Centralna ewidencja dokumentów handlowych

Główny moduł systemu umożliwia kompleksową obsługę dokumentów:

- dodawanie dokumentów wraz z załącznikami PDF,
- edycję i korektę metadanych,
- dwuetapowe usuwanie (anulowanie → trwałe usunięcie),
- pełną kontrolę widoczności danych w zależności od roli użytkownika.

Dokumenty anulowane:
- są niewidoczne dla użytkowników standardowych,
- są widoczne dla administratorów (oznaczone jako przekreślone).

---

### 2. Zaawansowane wyszukiwanie i filtrowanie

System udostępnia rozbudowany mechanizm filtrowania danych:

- filtrowanie po typie dokumentu,
- zakresy dat (od / do / konkretny dzień),
- filtrowanie wielokryterialne (sumowanie warunków),
- obsługa klawisza **Enter** jako wyzwalacza filtracji.

Dane prezentowane są w siatce z:
- możliwością kopiowania pojedynczych wartości,
- eksportem danych do **Excel (bez Office Automation)**.

---

### 3. Obsługa załączników i podgląd dokumentów

Każdy dokument może posiadać załącznik PDF (np. eksport z SAP).

**Wbudowany podgląd PDF umożliwia:**
- nawigację po stronach,
- powiększanie i pomniejszanie,
- zapis pliku na dysku,
- druk na drukarce domyślnej,
- wysyłkę dokumentu jako załącznik e-mail (Outlook).

---

### 4. Masowy import dokumentów (Administrator)

Moduł importu umożliwia:
- wsadowy import dokumentów PDF z folderów lokalnych lub sieciowych,
- automatyczne rozpoznanie typu dokumentu,
- raportowanie przebiegu importu (log + postęp),
- opcjonalne usuwanie plików po poprawnym imporcie,
- zapis raportu importu do pliku.

Proces może być:
- wstrzymany w dowolnym momencie,
- wznowiony z zachowaniem historii operacji.

---

### 5. Zarządzanie użytkownikami i uprawnieniami

Administrator systemu może:
- zarządzać uprawnieniami operacyjnymi (odczyt / dodawanie / edycja / usuwanie),
- definiować dodatkowe uprawnienia (druk, zapis, wysyłka PDF),
- wysyłać komunikaty do wybranych użytkowników bezpośrednio z aplikacji.

---

## 🛠 Warstwa techniczna

### Architektura

- aplikacja desktopowa **C# / WinForms**,
- architektura wielowarstwowa,
- centralna baza danych **Microsoft SQL Server**,
- deployment poprzez **Microsoft ClickOnce**.

### Baza danych (SQL Server)

- kontrola integralności danych,
- logiczny podział tabel (użytkownicy, dokumenty, słowniki),
- audyt operacji (anulowanie / usuwanie),
- wydajna obsługa dużych wolumenów danych.

[🔙 Powrót do README](../../../README.md)