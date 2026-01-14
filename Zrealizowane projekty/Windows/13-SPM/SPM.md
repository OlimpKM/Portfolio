# SPM - Security private managment (projekt prywatny)
Oprogramowanie napisane do elastycznego bezpiecznego zarządzania informacją oraz zadaniami.

![Status](https://img.shields.io/badge/Status-Na%20własny%20użytek-informational)  
![Technologia](https://img.shields.io/badge/Technologia-C%23%20%2F%20WinForms%20%2F%20SQLDatabase%2ENET-blue)  
![Typ](https://img.shields.io/badge/Typ-Projekt%20prywatny-lightgrey)  
![Rola](https://img.shields.io/badge/Rola-Autor%20%2F%20Developer-brightgreen)  
![Prawa autorskie](https://img.shields.io/badge/Prawa%20autorskie-Własność%20autora-lightgrey)

![SPM - kategoria](Screenshots/Kategoria%20-%20samochód.png)

> 📂 **Zasoby:** 
>  Wersja portable (latest) [SMP (portable).zip](SMP%20(portable).zip)
> Dokumentacja: [SPM.pdf](SPM.pdf)  
> Fragment zródeł: [C#](Samples/scr/) 


## Opis skrócony
Aplikacja umożliwia zarządzanie informacją w ramach kategorii oraz wspomaga realizację zadań.  

W ramach kategorii zarządzamy:
- danymi (prosta ewidencja klucz – wartość – uwagi - dokumenty),
- kontaktami (możliwość eksportu/importu do kontaktów google w formacie vCard),
- zdarzeniami,
- dokumentami,
- listą kont bankowych,
- rozrachunkami,
- przypomnieniami,
- ustawieniami / hasłami (połączony z mechanizmem automatycznego wprowadzania danych: użytkownik – hasło itp. po uruchomieniu wyzwalacza np. kombinacja klawiszy)

W ramach każdej kategorii dostępny jest folder w którym możemy przechowywać dokumenty. Zintegrowany eksplorator Windows umożliwia zarządzanie plikami i folderami. Został on rozszerzony o funkcje szyfrowania/ deszyfrowania plików. Funkcjonalność drag & drop umożliwia kopiowanie plików do innych aplikacji. Każda kategoria umożliwia stworzenie notatek w edytorze wizualnym (WYSIWYG). Otwierają się one w zakładkach danej kategorii.  

W ramach ewidencji zadań zarządzamy zadaniami. Każde zadanie opisują atrybuty (temat, data realizacji,oznaczenie, status itp.). Do każdego zadania tworzony jest folder. Zintegrowany eksplorator Windows umożliwia zarządzanie plikami i folderami. Dostępny jest edytor wizualny (WYSIWYG) który umożliwia tworzenie notatek. Zadania mogą być połączone z zewnętrznym systemem webowym poprzez zewnętrzny numer. Okno przeglądarki (Internet Explorer) będzie dostępne w ramach zakładek zadania. W ramach ewidencji zadań można przeszukiwać wszystkie zadania pod kątem wystąpienia danego ciągu znaków.

Dodatkowo aplikacja posiada możliwość stworzenia własnego
- menu ulubione (budujemy strukturę katalogów i umieszczamy pliki typu lnk. Po naciśnięciu danego linku otwiera się skojarzona aplikacja 32/64 bity)
- wzorce (budujemy strukturę katalogów i umieszczamy pliki wzorców, skrypty. Wzorce dostępne są w zintegrowanym eksploratorze Windows. Ze wzorców można kopiować
poszczególne pliki lub całe foldery
- symbole (budujemy strukturę katalogów i umieszczamy pliki tekstowe symboli. Zawartość będzie stanowiła treść do wklejenia do tworzenia notatek czy historii działań)
Dodatkowo aplikacja została wyposażona w mechanizm importu / eksportu na dysk Google ustawień oraz kontaktów.

Ciekawostki techniczne
- Szyfrowana kluczem 256 lokalna baza danych typu SQLight z własną obsługą ORM (mapowanie obiektowo-relacyjne),
- Rozbudowane standardowe kontrolki wizualne,
- Integracja z eksploratorem Windows,
- Integracja z Internet eksploratorem,
- Edytor wizualny (WYSIWYG) zbudowany w oparciu o darmową wersję Tiny MCE,
- Eksport / import kontaktów w formacie vCard,
- Integracja po API z usługami Google driver,
- Automatyczne wpisywanie tekstu bieżącej ewidencji (typer) to uruchomieniu wyzwalacza (kombinacja klawiszy) gdy aplikacja nie jest pierwszoplanowa lub naciśnięcie
klawisza i czas w celu przeniesienia aktywności,
- Aplikacja typu portale (przenaszalna)

[🔙 Powrót do README](../../../README.md)