# UmowySp – Bramka SMS
**Enterprise SMS Gateway & Message Delivery Management System**

![Status](https://img.shields.io/badge/Status-Zrealizowany-success)  
![Technology](https://img.shields.io/badge/Technologia-C%23%20%2F%20SQL%20Server-blue)  
![Platform](https://img.shields.io/badge/Platforma-Windows-lightgrey)  
![Auth](https://img.shields.io/badge/Autoryzacja-Active%20Directory-brightgreen)  
![Legal](https://img.shields.io/badge/Typ-Projekt%20komercyjny%20(UoP)-orange)  
![Owner](https://img.shields.io/badge/Prawa%20autorskie-Distribev-lightgrey)  

**Bramka SMS** to wewnętrzny system, zaprojektowany do **automatyzacji, kontroli kosztów oraz monitorowania masowej wysyłki wiadomości SMS** w organizacji.
Rozwiązanie pełni rolę centralnej bramki komunikacyjnej pomiędzy systemami biznesowymi (np. windykacyjnymi) a zewnętrznym dostawcą usług SMS.

System został zaprojektowany z naciskiem na:
- bezpieczeństwo,
- skalowalność,
- audytowalność,
- kontrolę kosztów wysyłki.

---

## 🔐 Autoryzacja i bezpieczeństwo

- Logowanie użytkowników realizowane poprzez **Active Directory (Windows Authentication)**
- Dostęp do funkcjonalności oparty o role i uprawnienia domenowe
- Klienci systemowi (aplikacje) identyfikowani poprzez **aktywne klucze API**
- Pełna rejestracja operacji (kto, kiedy, z jakiego źródła)

---

## 🧩 Architektura rozwiązania

System składa się z kilku logicznych komponentów:

- **Usługa backgroundowa**
  Odpowiada za:
  - komunikację z API dostawcy SMS,
  - wysyłkę wiadomości,
  - cykliczne pobieranie statusów doręczeń.

- **Aplikacja desktopowa (C# / Windows)**
  Narzędzie operacyjne do:
  - monitorowania wysyłek,
  - zarządzania blokadami (ręczne blokady numerów, historia blokad)
  - konfiguracji kluczy, limitów i wzorców treści (aktywność klucza, limity dzienne i miesięczne, godziny wysyłki, priorytety, statystyki dzienne i miesięczne),
  - zarządzanie wzorcami treśći sms(ów),
  - zlecenie wysyłki (możliwość użycia wzorców treści)

---

## 📡 Integracja z dostawcą SMS

Wysyłka realizowana była poprzez **SMSAPI** przy użyciu dostarczonego API.

### Mechanizm potwierdzeń (Delivery Reports)

System implementował **wielostopniowy retry policy** dla pobierania statusów doręczeń. Każda próba była rejestrowana wraz z licznikiem oraz planowaną kolejną próbą.

### Kontrola kosztów

- Koszt SMS liczony w punktach (zależny od długości wiadomości)
- Znaki specjalne liczone jako **2 znaki standardowe**
- Limity dzienne i miesięczne per klucz nadawcy
- Blokady numerów w celu eliminacji zbędnych kosztów

## 🗄 Warstwa danych

**MS SQL Server**

- procedury składowane,
- funkcje bazodanowe,
- transakcyjność,
- spójność danych pomiędzy usługą a UI.

Zlecenie wysyłki SMS(ów) może być rejestrowane:
- bezpośrednio w bazie (procedura składowana),
- przez aplikację desktopową,
- przez serwis WCF.

---

## 📊 Schemat procesu wysyłki SMS

```mermaid
graph TD
    A[System kliencki]
    B[Rejestracja przez procedurę DB]
    C[Rejestracja przez usługę WCF]

    D[Rejestracja SMS w bazie]
    E[Usługa backgroundowa]
    F[API SMSAPI]
    G[Pobranie statusu]
    H[Aktualizacja bazy]

    A --> B
    A --> C
    B --> D
    C --> D
    D --> E
    E --> F
    F --> G
    G --> H
```
[🔙 Powrót do README](../../../README.md)