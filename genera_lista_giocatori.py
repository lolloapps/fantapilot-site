import pandas as pd
from pathlib import Path
import argparse


def genera_lista_giocatori(file_statistiche, file_quotazioni, output):
    """
    Genera il file ListaGiocatori a partire dai file:
      - Statistiche Fantacalcio
      - Quotazioni Fantacalcio

    Viene utilizzato esclusivamente il foglio 'Tutti'.
    La chiave di unione è 'Id'.
    """

    # 1. Lettura del solo foglio "Tutti".
    # Nei file originali la prima riga è il titolo,
    # mentre le intestazioni reali sono nella seconda riga.
    statistiche = pd.read_excel(
        file_statistiche,
        sheet_name="Tutti",
        header=1,
        engine="openpyxl"
    )

    quotazioni = pd.read_excel(
        file_quotazioni,
        sheet_name="Tutti",
        header=1,
        engine="openpyxl"
    )

    # 2. Controllo delle colonne necessarie
    colonne_statistiche = ["Id", "R", "Pv", "Mv", "Fm"]
    colonne_quotazioni = ["Id", "Nome", "Squadra", "Qt.A", "FVM"]

    mancanti_stat = [
        c for c in colonne_statistiche
        if c not in statistiche.columns
    ]

    mancanti_quot = [
        c for c in colonne_quotazioni
        if c not in quotazioni.columns
    ]

    if mancanti_stat:
        raise ValueError(
            f"Colonne mancanti nel file statistiche: {mancanti_stat}"
        )

    if mancanti_quot:
        raise ValueError(
            f"Colonne mancanti nel file quotazioni: {mancanti_quot}"
        )

    # 3. Prendiamo solo le colonne che ci servono
    statistiche = statistiche[
        ["Id", "R", "Pv", "Mv", "Fm"]
    ]

    quotazioni = quotazioni[
        ["Id", "Nome", "Squadra", "Qt.A", "FVM"]
    ]

    # 4. Merge.
    # INNER JOIN equivale, nel nostro caso, a eliminare
    # le righe che nel precedente procedimento avevano Nome_quot vuoto.
    risultato = pd.merge(
        statistiche,
        quotazioni,
        on="Id",
        how="inner"
    )

    # 5. Se per sicurezza esistessero record senza Nome,
    # vengono esclusi.
    risultato = risultato[
        risultato["Nome"].notna()
        & risultato["Nome"].astype(str).str.strip().ne("")
    ].copy()

    # 6. Rinomina delle colonne
    risultato = risultato.rename(
        columns={
            "R": "ruolo",
            "Pv": "pgv",
            "Mv": "mv",
            "Fm": "fm",
            "Nome": "nome",
            "Squadra": "squadra",
            "Qt.A": "quotazione_base",
            "FVM": "fv",
        }
    )

    # 7. Ordine esatto delle colonne finali
    risultato = risultato[
        [
            "Id",
            "ruolo",
            "pgv",
            "mv",
            "fm",
            "nome",
            "squadra",
            "quotazione_base",
            "fv",
        ]
    ]

    # 8. Creazione Excel
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        risultato.to_excel(
            writer,
            sheet_name="Tutti",
            index=False
        )

        workbook = writer.book
        worksheet = writer.sheets["Tutti"]

        # Intestazione come nel file generato
        formato_header = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E78"
        })

        for colonna, nome in enumerate(risultato.columns):
            worksheet.write(
                0,
                colonna,
                nome,
                formato_header
            )

            valori = (
                [str(nome)]
                + risultato[nome]
                    .dropna()
                    .astype(str)
                    .head(200)
                    .tolist()
            )

            larghezza = min(
                max(len(v) for v in valori) + 2,
                24
            )

            worksheet.set_column(
                colonna,
                colonna,
                larghezza
            )

        # Prima riga bloccata
        worksheet.freeze_panes(1, 0)

        # Filtro automatico
        worksheet.autofilter(
            0,
            0,
            len(risultato),
            len(risultato.columns) - 1
        )

    print(f"File generato: {output}")
    print(f"Giocatori: {len(risultato)}")


if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Generazione Lista Giocatori Fantacalcio"
    )

    parser.add_argument(
        "statistiche",
        help="File Excel delle statistiche"
    )

    parser.add_argument(
        "quotazioni",
        help="File Excel delle quotazioni"
    )

    parser.add_argument(
        "-o",
        "--output",
        default="ListaGiocatori.xlsx",
        help="Nome del file Excel di output"
    )

    args = parser.parse_args()

    genera_lista_giocatori(
        args.statistiche,
        args.quotazioni,
        args.output
    )