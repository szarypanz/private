"""
Created on Mon Jul 13 19:37:44 2020
FOR OWCA
"""

# import paczek potrzebnych do działania kodu
import os
import sys
import pandas as pd
import openpyxl as pyxl


# ustalenie lokacji roboczej - skrypt należy uruchomić komendą:
# python C:\PRZYKLAD_SCIEZKI\script.py
# końcowe raporty będą zapisywane w tej samej lokacji
os.chdir(os.path.dirname(sys.argv[0]))


## do testowania pliku z poziomu IDE należy stosować poniższe 3 linijki:
#PATH = "C:\\Users\\twoj_folder_z_pythonem"
#os.chdir(PATH)
#del PATH


# zaczynamy od podania liczby tabel do przeskanowania:
LICZBA_TABEL = int(input("   Podaj liczbę tabel EM zawartych w pliku: "))

# ...pobrania informacji nt tabeli absorbancji...:
TABELA_ABSORBANCJI = input("   Czy plik zawiera tabelę absorbancji? (T/N): ")

# ...i podania nazwy pliku do wczytania:
ADRES_EXCELA = input("   Podaj nazwę pliku, razem z rozszerzeniem xls/xlsx: ")
print("      wczytywanie plików...")
GLOWNA_BAZA = pd.read_excel(ADRES_EXCELA)


# zmieniamy nazwy kolumn w Glownej Bazie na bardziej przejrzyste wartosci
KOLUMNY = []
X = 0
for i in GLOWNA_BAZA.columns:
    KOLUMNY.append("Kolumna "+str(X))
    X = X + 1
GLOWNA_BAZA.columns = KOLUMNY
del KOLUMNY


###############################################################################


print("      ustalanie listy tabel i ich adresów...")
# na podstawie informacji z inputu pętlą for tworzymy listę tabel:
LISTA_TABEL = [i for i in range(LICZBA_TABEL)]
for y in LISTA_TABEL:
    TABELA = "Read " + str(y + 1) + ":EM Spectrum"
    LISTA_TABEL[y] = TABELA


# tworzymy listę adresów tabel na podstawie ich nazw
WYCINEK = GLOWNA_BAZA.iloc[:, 0:2]
WYCINEK_INDEX = pd.Index(list(WYCINEK["Kolumna 0"]))


# Ustalamy miejsca startowe, tworząc ramkę danych z adresami i wypełniając ją
# pętlami FOR:
USTAWIENIA = pd.DataFrame(columns=["Start", "Stop","Start2", "Stop2"])
X = 0
for i in WYCINEK["Kolumna 0"]:
    if i in LISTA_TABEL:
        ADRES = WYCINEK_INDEX.get_loc(i) + 3
        USTAWIENIA.at[X, "Start"] = ADRES
        X = X + 1

USTAWIENIA["Start"] = USTAWIENIA["Start"].astype(int)


# Ustalamy koniec występowania danych dla kazdej tabeli:
X = 0
for i in LISTA_TABEL:
    LISTA_FAL = pd.Series(WYCINEK.iloc[USTAWIENIA.loc[X, "Start"]:, 1])
    LISTA_FAL = LISTA_FAL.fillna("KONIEC")
    USTAWIENIA.at[X, "Stop"] = LISTA_FAL[LISTA_FAL == "KONIEC"].index[0]
    X = X + 1

USTAWIENIA["Stop"] = USTAWIENIA["Stop"].astype(int)


# jesli plik zawiera tab absorbancji, sciągamy jej adres na podstawie warunku:
if TABELA_ABSORBANCJI in ("T", "t"):
    TABELA = "Read "+str(len(LISTA_TABEL) + 1)+":Spectrum"
    for i in WYCINEK["Kolumna 0"]:
        if i == TABELA:
            ADRES = WYCINEK_INDEX.get_loc(i) + 3
            USTAWIENIA.at[0, "Start2"] = ADRES
            USTAWIENIA["Start2"].fillna(0, inplace=True)
            USTAWIENIA["Start2"] = USTAWIENIA["Start2"].astype(int)

    LISTA_FAL = pd.Series(WYCINEK.iloc[USTAWIENIA.loc[0, "Start2"]:, 1])
    LISTA_FAL = LISTA_FAL.fillna("KONIEC")
    USTAWIENIA.at[0, "Stop2"] = LISTA_FAL[LISTA_FAL == "KONIEC"].index[0]
    USTAWIENIA["Stop2"].fillna(0, inplace=True)
    USTAWIENIA["Stop2"] = USTAWIENIA["Stop2"].astype(int)

del LISTA_FAL, WYCINEK, WYCINEK_INDEX


# Ustawiamy nowe nazwy tabel które ułatwią późniejsze sortowanie
X = 1
for i in LISTA_TABEL:
    if X < 10:
        LISTA_TABEL[X - 1] = "0" + str(X) + ": EM Spectrum"
    else:
        LISTA_TABEL[X - 1] = str(X) + ": EM Spectrum"
    X = X + 1


# tworzymy listę dołków na bazie adresu pierwszej tabelki z pliku Ustawień:
LISTA_DOLKOW = GLOWNA_BAZA.iloc[USTAWIENIA.loc[0, "Start"]-1, 2:]
LISTA_DOLKOW = [i for i in LISTA_DOLKOW if str(i) != "nan"]


###############################################################################


print("      aplikowanie ustawień...")
# wypełniamy listę wszystkich Długosci Fal:
# ustalamy w którym miejscu kończy się ostatnia tabelka
KONIEC = USTAWIENIA["Stop"]
KONIEC.dropna(inplace=True)
KONIEC = KONIEC.to_frame()
KONIEC = KONIEC.astype(int)


# na podstawie powyższego sciągamy wszystkie wartosci i konwertujemy je na INT:
WAVELENGHTS = GLOWNA_BAZA.iloc[:KONIEC["Stop"].iloc[-1], 1]
del KONIEC
WAVELENGHTS = WAVELENGHTS.to_frame()
WAVELENGHTS.columns = ['Wavelength']
WAVELENGHTS["Wavelength"] = pd.to_numeric(WAVELENGHTS["Wavelength"],
                                          downcast="integer",
                                          errors="coerce")
WAVELENGHTS.dropna(inplace=True)
WAVELENGHTS.drop_duplicates(inplace=True)
WAVELENGHTS = WAVELENGHTS.astype(int)


# poniższa linijka kodu wyrzuci wszystkie wartosci mniejsze niż 999
# - w wypadku innych wartosci należy podmienic wartosc w nawiasie
WAVELENGHTS = WAVELENGHTS[WAVELENGHTS["Wavelength"] <= 999]
WAVELENGHTS.sort_values(by=["Wavelength"], inplace=True)


###############################################################################


print("      edycja excela...")
# Do zapisania Excela potrzebne jest napisanie tzw Writera czyli narzędzia
# do edycji arkuszy
PLIK_EXCEL = pyxl.load_workbook(ADRES_EXCELA)


# po ustaleniu scieżki pliku przypisujemy go writerowi
WRITER = pd.ExcelWriter(ADRES_EXCELA, engine="openpyxl")
WRITER.book = PLIK_EXCEL


# powyższe narzędzie wykorzystamy do przenoszenia danych w kolejnych krokach,
# idąc pętlą FOR po kolejnych płytkach


###############################################################################


print("      tworzenie macierzy...")
# Przystępujemy do tworzenia własciwej macierzy - najpierw tworząc pustą ramkę:
MACIERZ = pd.DataFrame()
NR_DOLKA = 2   # Adres pod ktorym znajduje się kolumna dołka A1


# Nastepnie sciągamy wszystkie wartosci per każdy dołek:
for i in LISTA_DOLKOW:

    WIERSZ_MACIERZY = pd.Series(i)

    X = 0
    for tabela in LISTA_TABEL:
        # kopiujemy dane z głównych tabel:
        WYCINEK = GLOWNA_BAZA.iloc[USTAWIENIA.iloc[X, 0]:USTAWIENIA.iloc[X, 1],
                                   NR_DOLKA]
        WIERSZ_MACIERZY = WIERSZ_MACIERZY.append(WYCINEK, ignore_index=True)
        X = X + 1


    # Jesli oprócz tabel EM wyniki zawierają też dodatkową tabelę bez EM możemy
    # skopiować ją na podstawie poniższego warunku:
    if TABELA_ABSORBANCJI in ("T", "t"):
        WYCINEK = GLOWNA_BAZA.iloc[USTAWIENIA.iloc[0, 2]:USTAWIENIA.iloc[0, 3],
                                   NR_DOLKA]
        WIERSZ_MACIERZY = WIERSZ_MACIERZY.append(WYCINEK, ignore_index=True)


    # Transponujemy dane z dołka do tabelki i przechodzimy do następnego
    MACIERZ = MACIERZ.append(WIERSZ_MACIERZY, ignore_index=True)
    NR_DOLKA = NR_DOLKA + 1


# Korzystamy ze stworzonego wczeniej Writera by zapisać gotową Macierz w Excelu
MACIERZ.to_excel(WRITER, sheet_name="Macierz", index=None, startcol=1)


###############################################################################


print("      tworzenie map F2D...")
# Przygotowujemy tabelę docelową i wstawiamy w nią listę fal;
# jednoczesnie zapisujemy jej kopię na użytek późniejszej pętli FOR:
TABELA_DOCELOWA = pd.DataFrame(columns=["Wavelength"])
TABELA_DOCELOWA = TABELA_DOCELOWA.append(WAVELENGHTS, sort=True)
TABELA_KOPIA = TABELA_DOCELOWA.copy()
del WAVELENGHTS


# na jej podstawie pętlami FOR można skopiować wszystkie wartosci:
# tworzymy tabelkę zawierającą wszystkie wartosci fal z danego dołka
NR_DOLKA = 2   # Adres pod ktorym znajduje się kolumna dołka A1

for dolek in LISTA_DOLKOW:

    # Używamy kopii tabeli żeby zachować oryginał na kolejne przejscia pętli
    TABELA_DOCELOWA = TABELA_KOPIA

    X = 0
    for i in LISTA_TABEL:
        # Tworzymy listę fal i wartosci dla danego dołka:
        FALA = GLOWNA_BAZA.iloc[USTAWIENIA.iloc[X, 0]:USTAWIENIA.iloc[X, 1], 1]
        DOLEK = GLOWNA_BAZA.iloc[USTAWIENIA.iloc[X, 0]:USTAWIENIA.iloc[X, 1],
                                 NR_DOLKA]
        WYCINEK = pd.concat([FALA, DOLEK], axis=1)
        WYCINEK.columns = ["Wavelength", i]


        # po czym dokonujemy pythonowego vlookupowania, po fali
        TABELA_DOCELOWA = pd.merge(TABELA_DOCELOWA, WYCINEK, on="Wavelength",
                                   how="left")
        TABELA_DOCELOWA = TABELA_DOCELOWA.fillna("")
        X = X + 1


    # Sortujemy wyniki od największej tabeli
    TABELA_DOCELOWA.sort_index(axis=1, ascending=False, inplace=True)


    # Korzystamy ze stworzonego wczeniej Writera by przeniesc Tabelę Docelową
    # na nowy arkusz
    TABELA_DOCELOWA.to_excel(WRITER, sheet_name=dolek, index=None)


    NR_DOLKA = NR_DOLKA + 1   # idziemy do kolejnej kolumny z pliku xls


del GLOWNA_BAZA


###############################################################################


# zapisujemy zmiany i zamykamy narzędzie
print("      zapisywanie...")
WRITER.save()


###############################################################################


print("           __  _")
print("       .-.'  `; `-._  __  _")
print("      (_,         .-:'  `; `-._")
print('    ,"o"(        (_,           )')
print('   (__,-'"'"'      ,"o"(            )>')
print("      (       (__,-'  GOTOWE    )")
print("       `-'._.--._(             )")
print("          |||  |||`-'._.--._.-'")
print("                     |||  |||")


###############################################################################

