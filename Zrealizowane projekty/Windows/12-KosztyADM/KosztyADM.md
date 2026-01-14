# KosztyADM
*System do zarządzania i analizy kosztów administracyjnych w strukturze oddziałowej.*

![Status](https://img.shields.io/badge/Status-Zrealizowany-success)  
![Technology](https://img.shields.io/badge/Technologia-C%23%20%2F%20MS%20SQL-blue)  
![Platform](https://img.shields.io/badge/Platforma-Windows%20(WinForms)-informational)  
![Legal](https://img.shields.io/badge/Typ-Projekt%20komercyjny%20(UoP)-orange)  
![Role](https://img.shields.io/badge/Rola-Full--Stack%20Developer-brightgreen)  
![Owner](https://img.shields.io/badge/Prawa%20autorskie-Distribev-lightgrey)

Oprogramowanie **KosztyADM** to wewnętrzny system klasy back-office, zaprojektowany do **ewidencji, kontroli oraz raportowania kosztów administracyjnych** ponoszonych przez poszczególne oddziały firmy. System wspiera procesy rozliczeniowe oraz zapewnia spójność danych finansowych na poziomie całej organizacji.

---

## 🚀 Kluczowe funkcjonalności

### 1. Moduł ewidencji podstawowej
Centralny moduł odpowiedzialny za utrzymanie słowników oraz danych referencyjnych wykorzystywanych w całym systemie.
* **Ewidencja oddziałów:** Struktura organizacyjna,
* **Ewidencja kontrahentów:** Dane dostawców usług i towarów generujących koszty administracyjne,
* **Ewidencja stawek:** Definicje kosztów cyklicznych i zmiennych (np. media, usługi, opłaty stałe),

### 2. Moduł rozliczania kosztów
Obsługa rzeczywistych kosztów administracyjnych z podziałem na oddziały i kategorie kosztowe,
* **Rejestr kosztów:** Przypisywanie kosztów do oddziału, kontrahenta i rodzaju kosztu,
* **Kontrola spójności danych:** Walidacja kompletności i poprawności wprowadzanych informacji,

### 3. Moduł raportowy
Wbudowany system raportowania wspierający analizę kosztów i przygotowanie zestawień zarządczych.

---

## 🛠 Warstwa techniczna i implementacja

System **KosztyADM** został wykonany jako aplikacja desktopowa w technologii **WinForms**, z naciskiem na stabilność, prostotę obsługi oraz integrację z istniejącą infrastrukturą IT.

### 1. Persistence Layer: MS SQL Server
* Relacyjna baza danych zapewniająca spójność i integralność danych finansowych,
* Optymalizacja zapytań,

### 2. Warstwa aplikacji (C# / WinForms)
* Klasyczna architektura warstwowa (UI → Logika biznesowa → Dane),
* Czytelne formularze ewidencyjne,

### 3. Bezpieczeństwo i dostęp
* **Uwierzytelnianie użytkowników:** Integracja z **Active Directory**.
* **Kontrola dostępu:** Oparcie uprawnień o konta domenowe użytkowników.
* **Brak lokalnych haseł:** Pełne wykorzystanie infrastruktury domenowej.

[🔙 Powrót do README](../../../README.md)