import pandas as pd

"""
TODO: 

Executable met interface van maken?

"""
global df
df = pd.read_excel('origineel\\Ledenbestand.xlsx')


def export_vrouwen():
    """
    Simpele versie van export_gegevens_logic_test.
    :return:
    """
    voornaam, tussenvoegsel, achternaam, straat, huisnummer, postcode, woonplaats = [], [], [], [], [], [], []
    for i in range(len(df)):
        if df['Vrouw'].iloc[i]:
            voornaam.append(df['voornaam'].iloc[i])
            tussenvoegsel.append(df['tussenvoegsel'].iloc[i])
            achternaam.append(df['achternaam'].iloc[i])
            straat.append(df['straat'].iloc[i])
            huisnummer.append(df['huisnummer'].iloc[i])
            postcode.append(df['postcode'].iloc[i])
            woonplaats.append(df['woonplaats'].iloc[i])

    df_angfo = pd.DataFrame(voornaam, columns=['voornaam'])
    df_angfo['tussenvoegsel'] = pd.Series(tussenvoegsel, index=df_angfo.index)
    df_angfo['achternaam'] = pd.Series(achternaam, index=df_angfo.index)
    df_angfo['straat'] = pd.Series(straat, index=df_angfo.index)
    df_angfo['huisnummer'] = pd.Series(huisnummer, index=df_angfo.index)
    df_angfo['postcode'] = pd.Series(postcode, index=df_angfo.index)
    df_angfo['woonplaats'] = pd.Series(woonplaats, index=df_angfo.index)
    df_angfo.to_excel('output\\vrouwen_leden.xlsx', 'vrouwen_leden')


def export_gegevens(te_doorzoeken_kolom, is_gelijk_aan, exporteer_gegevens, bestandsnaam):
    """
    Dit is een algemene functie welke kijkt welke leden in het ingeladen ledenbestand voldoen aan de logic check.
    Wanneer het betreffende lid voldoet dan worden de gegevens zoals deze in de lijst 'kolomen' staan opgeslagen in een
    nieuw excel bestand.
    De tekst in 'logic_check' en 'kolomen' dienen exact overeen te komen met de betreffende kolomen in het originele
    excel bestand.
    :param logic_check: Een string van de naam van de header waar de logic_check op word uitgevoerd.
    :param gegevens: Een array met de namen van de headers welke geexporteerd dienen te worden
    :param bestandsnaam: De naam van het te exporteren bestand
    :return:
    """
    empt_lists = []
    # array met lege lijsten maken
    for n in range(len(exporteer_gegevens)):
        empt_lists.append([])

    # gegevens aan de gemaakte array toevoegen in de betreffende lijsten
    for i in range(len(df)):
        if df[te_doorzoeken_kolom].iloc[i] == is_gelijk_aan:
            for j in range(len(exporteer_gegevens)):
                empt_lists[j].append(df[exporteer_gegevens[j]].iloc[i])

    df_angfo = pd.DataFrame(empt_lists[0], columns=[exporteer_gegevens[0]])

    # geheel omzetten naar een pandas dataframe
    for k in range(len(exporteer_gegevens)):
        df_angfo[exporteer_gegevens[k]] = pd.Series(empt_lists[k], index=df_angfo.index)

    df_angfo.to_excel(bestandsnaam, 'Sheet1')
    print('Het bestand is opgeslagen als: %s' % bestandsnaam)


def nieuw_ledenbestand():
    """
    Per row in df loop schrijven die door gehele df_a gaat om op een match te controleren.

    TODO

    Zoek voornaam van df_a in df, noteer indices in lijst1 van df waar dit voorkomt.
    Zoek tussenvoegsel van df_a in df, noteer indices in lijst2 van df waar dit voorkomt.
    Zoek achterneem van df_a, noteer indices in lijst3 van df waar dit voorkomt.

    Bevatten de drie lijsten allemaal een keer hetzelfde getal (index waarde van df)?
    Zo ja, dan is er een match.
    :return:
    """
    """
    df_a = pd.read_excel('origineel\\Angfo.xlsx')
    df_a_naam = pd.DataFrame()
    global df
    df = df.fillna('')

    # (len(df_a['achternaam'].str.contains(df['achternaam'][i], na=False)) == 0)

    #  Arrays met complete namen maken
    naam_df = []
    naam_df_a = []
    for i in range(len(df)):
        naam_df.append(str(df['voornaam'][i])+' '+str(df['tussenvoegsel'][i])+' '+str(df['achternaam'][i]))
    for i in range(len(df_a)):
        naam_df_a.append(str(df_a['voornaam'][i])+' '+str(df_a['tussenvoegsel'][i])+' '+str(df_a['achternaam'][i]))

    df_namen = pd.DataFrame(naam_df, columns=['naam'])
    df_a_namen = pd.DataFrame(naam_df_a, columns=['naam'])

    x = df_a_namen['naam'].str.contains(df_namen['naam'][0])
    print(x)
    """

def nieuw_ledenbestand():
    """
    Zoek voornaam van df_a in df, noteer indices in lijst1 van df waar dit voorkomt.
    Zoek tussenvoegsel van df_a in df, noteer indices in lijst2 van df waar dit voorkomt.
    Zoek achterneem van df_a, noteer indices in lijst3 van df waar dit voorkomt.

    Bevatten de drie lijsten allemaal een keer hetzelfde getal (index waarde van df)?
    Zo ja, dan is er een match.
    :return:
    """
    df_a = pd.read_excel('origineel\\Angfo.xlsx')
    to_drop = df_a.columns.difference(['voornaam', 'achternaam', 'tussenvoegsel'])
    df_a.drop(to_drop, 1, inplace=True)  # drop alle columns behalve degene in to_drop
    print(df_a)

#export_gegevens('Vrouw', True,
#                ['voornaam', 'tussenvoegsel', 'achternaam', 'straat', 'huisnummer', 'postcode', 'woonplaats'],
#                'output\\vrouwen_leden.xlsx')

#export_gegevens('lid status', 'erelid',
#                ['voornaam', 'tussenvoegsel', 'achternaam', 'straat', 'huisnummer', 'postcode', 'woonplaats', 'geboortedatum'],
#                'output\\ere_leden.xlsx')

nieuw_ledenbestand()