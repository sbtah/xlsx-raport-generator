import pandas as pd
import os


os.chdir('Raport')
FILE_PATH = 'Raport_All.xlsx'
 
df = pd.read_excel(FILE_PATH, index_col="Nazwa")

marki = df["Marka"]
os.mkdir('Wygenerowane')

marki_lista = set([marka for marka in marki])
for marka in marki_lista:

    df['Średnia sprzedaż w Miesiącu'] = df['Ilość sprzedana'] / 12
    df['Średni zapas Na 6M'] = df['Średnia sprzedaż w Miesiącu'] * 6
    df['Stan-Zapas'] = df['Stan teoretyczny'] - \
        df['Średni zapas Na 6M']
    df['Ilość do Zwrotu'] = df['Stan teoretyczny'] - \
        (df['Średnia sprzedaż w Miesiącu'] * 3)
    df['Ilość do Zwrotu Zaokrąglone'] = df['Ilość do Zwrotu'].round()

    df.sort_index(inplace=True)

    filt = (df['Stan-Zapas'] >= 1) & (df['Marka']
                                            == marka)  # this will create many filters

    raport2 = df.loc[filt, ['EAN Prio', 'Ilość do Zwrotu Zaokrąglone']]

    writer = pd.ExcelWriter(f"Wygenerowane/Raport_{marka}.xlsx")
    df[filt].to_excel(writer, sheet_name="Raport")
    raport2.to_excel(writer, sheet_name="Raport EAN i Ilość")

    print(f"Raport_{marka} - WYKONANY!")

    writer.save()