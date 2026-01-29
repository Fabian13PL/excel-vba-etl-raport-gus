# Automatyczny raport wynagrodzeń – Excel VBA

## Opis projektu
Projekt przedstawia kompletny proces automatycznego raportowania danych statystycznych w Excelu z wykorzystaniem **VBA**.  
Makro importuje dane CSV (UTF-8), przekształca je do postaci analitycznej (*long format*), buduje tabelę przestawną oraz generuje wykres.

Projekt został wykonany jako element **portfolio analityczno-automatyzacyjnego**.

---

## Zakres funkcjonalny
- Import danych CSV z uwzględnieniem kodowania **UTF-8**
- Czyszczenie i normalizacja danych
- Transformacja danych z układu szerokiego do **DANE_LONG**
- Automatyczne tworzenie:
  - tabeli przestawnej
  - wykresu liniowego
- Wybór analizowanego wskaźnika z poziomu arkusza START
- Pełna automatyzacja procesu jednym przyciskiem

---

## Struktura arkuszy
- **START** – panel sterujący (wybór wskaźnika, status)
- **DANE_RAW** – surowe dane z CSV
- **DANE_LONG** – dane w formacie analitycznym
- **PIVOT** – tabela przestawna
- **RAPORT** – wykres końcowy

---

## Technologie
- Microsoft Excel
- VBA (Visual Basic for Applications)
- Tabele przestawne
- Przetwarzanie danych tekstowych (CSV, UTF-8)

---

## Źródło danych
Dane pochodzą z **publicznej bazy GUS**  
Zakres: wynagrodzenia i świadczenia społeczne – dane miesięczne.

Plik CSV jest wykorzystywany jako **wejście do procesu ETL** realizowanego w VBA.

---

## Jak uruchomić
1. Otwórz plik `.xlsm`
2. W arkuszu **START**:
   - wybierz wskaźnik
   - kliknij „Import”
   - kliknij „Generuj raport”
3. Raport i wykres zostaną wygenerowane automatycznie

---

## Cel projektu
Celem projektu było pokazanie umiejętności:
- pracy na rzeczywistych danych publicznych
- automatyzacji raportów w Excelu
- projektowania prostego procesu ETL
- przygotowania danych pod analizę i wizualizację
Projekt symuluje zadanie analityka polegające na automatycznym raportowaniu danych cyklicznych.
