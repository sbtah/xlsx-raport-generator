import pandas as pd
import os

#For clean use create a 'Raport' folder in same dir that script is located and copy original file there.
os.chdir('Raport')
FILE_PATH = 'Raport_All.xlsx'
 
df = pd.read_excel(FILE_PATH, index_col="Nazwa")

marki = df["Marka"]
os.mkdir('Wygenerowane')

marki_lista = set([marka for marka in marki])
for marka in marki_lista:

    # Theese are the columns I need to generate to calculate which articles we would like to return to producer. In previous years I generated those files manually in excel that took my arround a month of work since I was doing this as en extra task.
    df['Średnia sprzedaż w Miesiącu'] = df['Ilość sprzedana'] / 12
    df['Średni zapas Na 6M'] = df['Średnia sprzedaż w Miesiącu'] * 6
    df['Stan-Zapas'] = df['Stan teoretyczny'] - \
        df['Średni zapas Na 6M']
    df['Ilość do Zwrotu'] = df['Stan teoretyczny'] - \
        (df['Średnia sprzedaż w Miesiącu'] * 3)
    df['Ilość do Zwrotu Zaokrąglone'] = df['Ilość do Zwrotu'].round()

    
    df.sort_index(inplace=True)

    # This is the main filter I use on generated columns.
    filt = (df['Stan-Zapas'] >= 1) & (df['Marka']
                                            == marka)  # this will create many filters

    # This is a 2nd spreadsheet generated for lazy managers that prepares returns ;)
    raport2 = df.loc[filt, ['EAN Prio', 'Ilość do Zwrotu Zaokrąglone']]

    # This part saves file with specified sheets.
    writer = pd.ExcelWriter(f"Wygenerowane/Raport_{marka}.xlsx")
    df[filt].to_excel(writer, sheet_name="Raport")
    raport2.to_excel(writer, sheet_name="Raport EAN i Ilość")

    # Some information about status of program.
    print(f"Raport_{marka} - WYKONANY!")

    writer.save()