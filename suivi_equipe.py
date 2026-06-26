"""
MODÈLE DE SUIVI D'ÉQUIPE – FOOTBALL JEUNES
Génère un fichier .xlsx compatible Numbers / Excel

USAGE (Mac) :
    pip3 install openpyxl
    python3 suivi_equipe.py
→ Ouvrir suivi_equipe.xlsx avec Numbers (double-clic)
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter, column_index_from_string

# ─── Couleurs ────────────────────────────────────────────────────────────────

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=11, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic)

def align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def border_thin():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def border_medium():
    s = Side(style="medium", color="1a237e")
    return Border(left=s, right=s, top=s, bottom=s)

# Palette mois
MONTH_FILLS = [
    "1565C0","6A1B9A","00695C","BF360C",
    "4527A0","AD1457","2E7D32","F57F17",
    "0277BD","558B2F"
]

MONTHS = ["SEPTEMBRE","OCTOBRE","NOVEMBRE","DÉCEMBRE",
          "JANVIER","FÉVRIER","MARS","AVRIL","MAI","JUIN"]

NB_PLAYERS   = 30
SESSIONS_PM  = 9    # séances par mois
NB_MATCHS_P3 = 30   # convocations
NB_MATCHS_P4 = 20   # feuille de match

# ─── Helpers ─────────────────────────────────────────────────────────────────

def col(n):
    """Numéro de colonne → lettre(s)"""
    return get_column_letter(n)

def set_cell(ws, row, column, value=None, formula=None,
             bg=None, bold=False, color="000000", size=11,
             halign="left", valign="center", wrap=False,
             num_format=None, italic=False, border=False):
    c = ws.cell(row=row, column=column)
    if formula:
        c.value = formula
    elif value is not None:
        c.value = value
    if bg:
        c.fill = fill(bg)
    c.font = font(bold=bold, color=color, size=size, italic=italic)
    c.alignment = align(halign, valign, wrap)
    if num_format:
        c.number_format = num_format
    if border:
        c.border = border_thin()
    return c

def merge(ws, r1, c1, r2, c2, value=None, formula=None,
          bg=None, bold=False, color="000000", size=11,
          halign="center", valign="center", wrap=False,
          num_format=None, italic=False):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = set_cell(ws, r1, c1, value=value, formula=formula,
                 bg=bg, bold=bold, color=color, size=size,
                 halign=halign, valign=valign, wrap=wrap,
                 num_format=num_format, italic=italic)
    return c

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 1 : LISTE DES JOUEURS
# ─────────────────────────────────────────────────────────────────────────────

def build_page1(ws):
    ws.title = "Liste Joueurs"

    # Largeurs
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 22

    # Titre
    merge(ws, 1,1,1,6, "📋 LISTE DES JOUEURS",
          bg="1A237E", bold=True, color="FFFFFF", size=14)
    ws.row_dimensions[1].height = 32

    # Sous-titre
    merge(ws, 2,1,2,6, "Saison 2025/2026 – FC [Nom du club] – Catégorie [U13/U15…]",
          bg="E8EAF6", size=11, italic=True, color="3949AB")
    ws.row_dimensions[2].height = 22

    # En-têtes
    headers = ["N°", "Nom", "Prénom", "Date de naissance", "Poste", "Contact parent"]
    for i, h in enumerate(headers, 1):
        set_cell(ws, 3, i, h, bg="3949AB", bold=True, color="FFFFFF",
                 halign="center", valign="center", border=True)
    ws.row_dimensions[3].height = 22

    # Lignes joueurs
    for r in range(4, 4 + NB_PLAYERS):
        bg_row = "E8EAF6" if r % 2 == 0 else "FFFFFF"
        set_cell(ws, r, 1, r - 3, bg=bg_row, halign="center", border=True)
        for c_idx in range(2, 7):
            set_cell(ws, r, c_idx, bg=bg_row, border=True)

    ws.freeze_panes = "A4"

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 2 : PRÉSENCE ENTRAÎNEMENTS
# ─────────────────────────────────────────────────────────────────────────────

def build_page2(ws):
    ws.title = "Présence Entraînement"

    COLS_PM = SESSIONS_PM + 2          # 9 séances + TOT + %
    FIRST_DATA_COL = 2
    PLAYER_START_ROW = 5

    # Colonne nom
    ws.column_dimensions["A"].width = 22

    # Dropdown P / AJ / AI / B
    dv = DataValidation(
        type="list",
        formula1='"P,AJ,AI,B"',
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Valeur invalide",
        error="Choisir : P, AJ, AI ou B"
    )
    ws.add_data_validation(dv)

    # ── Ligne 1 : Titre ──────────────────────────────────────────────────────
    total_cols = 1 + len(MONTHS) * COLS_PM
    merge(ws, 1,1,1,total_cols,
          "📅 SUIVI DE PRÉSENCE AUX ENTRAÎNEMENTS – Saison 2025/2026",
          bg="1A237E", bold=True, color="FFFFFF", size=13)
    ws.row_dimensions[1].height = 30

    # ── Ligne 2 : Légende ────────────────────────────────────────────────────
    legend = [
        (1,"A",  "Légende :", "FFFFFF", False),
        (2,"B",  "P = Présent",              "C8E6C9", True),
        (4,"D",  "AJ = Absence Justifiée",   "FFF9C4", True),
        (6,"F",  "AI = Absence Injustifiée", "FFCDD2", True),
        (8,"H",  "B = Blessé / Malade",      "E1BEE7", True),
    ]
    for col_n, _, txt, bg, brd in legend:
        set_cell(ws, 2, col_n, txt, bg=bg, bold=True,
                 halign="center", border=brd)

    # ── Ligne 3-4 : En-têtes mois & séances ──────────────────────────────────
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 18

    # Cellule A3-A4 : label
    merge(ws, 3,1,4,1, "NOM / PRÉNOM",
          bg="37474F", bold=True, color="FFFFFF")

    col_cursor = FIRST_DATA_COL
    for m_idx, month in enumerate(MONTHS):
        month_start = col_cursor
        bg_month = MONTH_FILLS[m_idx % len(MONTH_FILLS)]

        # Mois (ligne 3, merge sur COLS_PM colonnes)
        merge(ws, 3, col_cursor, 3, col_cursor + COLS_PM - 1,
              month, bg=bg_month, bold=True, color="FFFFFF")

        # Séances (ligne 4)
        for s in range(1, SESSIONS_PM + 1):
            ws.column_dimensions[col(col_cursor)].width = 8
            set_cell(ws, 4, col_cursor, f"S{s}",
                     bg="E3F2FD", bold=True, halign="center", size=8)
            col_cursor += 1

        # TOT
        ws.column_dimensions[col(col_cursor)].width = 6
        set_cell(ws, 4, col_cursor, "TOT",
                 bg="BBDEFB", bold=True, halign="center", size=8)
        col_cursor += 1

        # %
        ws.column_dimensions[col(col_cursor)].width = 6
        set_cell(ws, 4, col_cursor, "%",
                 bg="BBDEFB", bold=True, halign="center", size=8)
        col_cursor += 1

    # ── Lignes joueurs ────────────────────────────────────────────────────────
    for r in range(PLAYER_START_ROW, PLAYER_START_ROW + NB_PLAYERS):
        player_idx = r - PLAYER_START_ROW + 2   # ligne dans Liste Joueurs (B2, C2…)
        bg_row = "ECEFF1" if r % 2 == 0 else "FFFFFF"

        # Nom (formule depuis Page 1)
        set_cell(ws, r, 1,
                 formula=f"='Liste Joueurs'!B{player_idx}&\" \"&'Liste Joueurs'!C{player_idx}",
                 bg="CFD8DC", bold=True)

        col_cursor = FIRST_DATA_COL
        for m_idx in range(len(MONTHS)):
            session_start_col = col_cursor

            # Cellules séances + dropdown
            for s in range(SESSIONS_PM):
                c = ws.cell(row=r, column=col_cursor)
                c.fill = fill(bg_row)
                c.alignment = align("center")
                dv.sqref = f"{dv.sqref} {col(col_cursor)}{r}" if dv.sqref else f"{col(col_cursor)}{r}"
                col_cursor += 1

            # TOT
            sc = col(session_start_col)
            ec = col(session_start_col + SESSIONS_PM - 1)
            set_cell(ws, r, col_cursor,
                     formula=f"=COUNTIF({sc}{r}:{ec}{r},\"P\")",
                     bg="BBDEFB", bold=True, halign="center")
            col_cursor += 1

            # %
            tot_cell = f"{col(col_cursor-1)}{r}"
            cnt_range = f"{col(session_start_col)}{r}:{col(session_start_col+SESSIONS_PM-1)}{r}"
            set_cell(ws, r, col_cursor,
                     formula=f"=IFERROR({tot_cell}/COUNTA({cnt_range}),\"\")",
                     bg="BBDEFB", bold=True, halign="center",
                     num_format="0%")
            col_cursor += 1

    # Ajout global du dropdown sur la plage de données séances
    session_end_col = 1 + len(MONTHS) * COLS_PM - len(MONTHS) * 2  # hors TOT/%
    # On re-crée un DV global plus simple
    dv2 = DataValidation(
        type="list",
        formula1='"P,AJ,AI,B"',
        showDropDown=False
    )
    # Range des séances uniquement (chaque mois, colonnes séances)
    ranges = []
    c2 = FIRST_DATA_COL
    for _ in MONTHS:
        ranges.append(f"{col(c2)}{PLAYER_START_ROW}:{col(c2+SESSIONS_PM-1)}{PLAYER_START_ROW+NB_PLAYERS-1}")
        c2 += COLS_PM
    dv2.sqref = " ".join(ranges)
    ws.add_data_validation(dv2)

    ws.freeze_panes = f"B{PLAYER_START_ROW}"

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 3 : CONVOCATIONS MATCHS
# ─────────────────────────────────────────────────────────────────────────────

def build_page3(ws):
    ws.title = "Convocations Matchs"

    PLAYER_START_ROW = 4

    ws.column_dimensions["A"].width = 22
    for m in range(1, NB_MATCHS_P3 + 1):
        ws.column_dimensions[col(m + 1)].width = 11

    # Titre
    merge(ws, 1,1,1, NB_MATCHS_P3 + 1,
          "📣 CONVOCATIONS MATCHS – Saison 2025/2026",
          bg="1A237E", bold=True, color="FFFFFF", size=13)
    ws.row_dimensions[1].height = 30

    # Légende
    set_cell(ws, 2, 1, "Légende :", bold=True)
    set_cell(ws, 2, 2, "✅ Convoqué",         bg="C8E6C9", bold=True, halign="center", border=True)
    set_cell(ws, 2, 4, "❌ Non convoqué",      bg="FFCDD2", bold=True, halign="center", border=True)
    set_cell(ws, 2, 6, "⚠️ Convoqué absent",   bg="FFF9C4", bold=True, halign="center", border=True)

    # En-têtes matchs (ligne 3)
    set_cell(ws, 3, 1, "NOM / PRÉNOM",
             bg="37474F", bold=True, color="FFFFFF", halign="center", border=True)
    for m in range(1, NB_MATCHS_P3 + 1):
        set_cell(ws, 3, m + 1, f"Match {m}",
                 bg="1565C0", bold=True, color="FFFFFF", halign="center", border=True)
    ws.row_dimensions[3].height = 20

    # Dropdown convocation
    dv = DataValidation(
        type="list",
        formula1='"✅ Convoqué,❌ Non convoqué,⚠️ Convoqué absent"',
        showDropDown=False
    )
    dv.sqref = f"B{PLAYER_START_ROW}:{col(NB_MATCHS_P3+1)}{PLAYER_START_ROW+NB_PLAYERS-1}"
    ws.add_data_validation(dv)

    # Lignes joueurs
    for r in range(PLAYER_START_ROW, PLAYER_START_ROW + NB_PLAYERS):
        player_idx = r - PLAYER_START_ROW + 2
        bg_row = "ECEFF1" if r % 2 == 0 else "FFFFFF"

        set_cell(ws, r, 1,
                 formula=f"='Liste Joueurs'!B{player_idx}&\" \"&'Liste Joueurs'!C{player_idx}",
                 bg="CFD8DC", bold=True)
        for m in range(1, NB_MATCHS_P3 + 1):
            ws.cell(row=r, column=m + 1).fill = fill(bg_row)
            ws.cell(row=r, column=m + 1).alignment = align("center")

    # Totaux
    tot_row = PLAYER_START_ROW + NB_PLAYERS
    pct_row = tot_row + 1
    set_cell(ws, tot_row, 1, "Nb convoqués", bg="37474F", bold=True, color="FFFFFF")
    set_cell(ws, pct_row, 1, "% convoqués",  bg="37474F", bold=True, color="FFFFFF")
    for m in range(1, NB_MATCHS_P3 + 1):
        c_letter = col(m + 1)
        set_cell(ws, tot_row, m + 1,
                 formula=f'=COUNTIF({c_letter}{PLAYER_START_ROW}:{c_letter}{PLAYER_START_ROW+NB_PLAYERS-1},"✅ Convoqué")',
                 bg="BBDEFB", bold=True, halign="center")
        set_cell(ws, pct_row, m + 1,
                 formula=f'=IFERROR({c_letter}{tot_row}/COUNTA({c_letter}{PLAYER_START_ROW}:{c_letter}{PLAYER_START_ROW+NB_PLAYERS-1}),"")',
                 bg="BBDEFB", bold=True, halign="center", num_format="0%")

    ws.freeze_panes = f"B{PLAYER_START_ROW}"

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 4 : FEUILLE DE MATCH
# ─────────────────────────────────────────────────────────────────────────────

def build_page4(ws):
    ws.title = "Feuille de Match"

    COLS_PM4   = 4   # Tit/Rem | Minutes | Buts | Passes décisives
    HEADER_ROWS = 7
    PLAYER_START_ROW = 8
    FIRST_MATCH_COL  = 2

    ws.column_dimensions["A"].width = 22
    ws.row_dimensions[1].height = 30

    # Titre
    total_cols = 1 + NB_MATCHS_P4 * COLS_PM4
    merge(ws, 1,1,1, total_cols,
          "⚽ FEUILLE DE MATCH – Saison 2025/2026",
          bg="1B5E20", bold=True, color="FFFFFF", size=13)

    # Labels colonne A (lignes 2-7)
    labels_a = {2:"MATCH", 3:"Date", 4:"Type", 5:"Dom/Ext", 6:"Score", 7:"NOM / PRÉNOM"}
    for row_n, lbl in labels_a.items():
        set_cell(ws, row_n, 1, lbl,
                 bg="263238", bold=True, color="FFFFFF",
                 halign="center", valign="center",
                 size=8 if row_n == 7 else 9)

    # Dropdowns
    dv_type = DataValidation(type="list", formula1='"Championnat,Amical,Coupe,Tournoi"', showDropDown=False)
    dv_domext = DataValidation(type="list", formula1='"Domicile,Extérieur,Terrain neutre"', showDropDown=False)
    dv_role = DataValidation(type="list", formula1='"Titulaire,Remplaçant,N/A"', showDropDown=False)
    ws.add_data_validation(dv_type)
    ws.add_data_validation(dv_domext)
    ws.add_data_validation(dv_role)

    MATCH_COLORS = ["1B5E20","1A237E","4A148C","B71C1C","E65100",
                    "006064","33691E","880E4F","1565C0","37474F"]

    for m in range(NB_MATCHS_P4):
        start_col = FIRST_MATCH_COL + m * COLS_PM4
        mc = MATCH_COLORS[m % len(MATCH_COLORS)]

        for sc in range(COLS_PM4):
            ws.column_dimensions[col(start_col + sc)].width = 9

        # Ligne 2 : "Match N"
        merge(ws, 2, start_col, 2, start_col + COLS_PM4 - 1,
              f"Match {m+1}", bg=mc, bold=True, color="FFFFFF")

        # Ligne 3 : Date
        dc = ws.cell(row=3, column=start_col)
        ws.merge_cells(start_row=3, start_column=start_col,
                       end_row=3, end_column=start_col + COLS_PM4 - 1)
        dc.fill = fill("E8F5E9")
        dc.number_format = "DD/MM/YYYY"
        dc.alignment = align("center")

        # Ligne 4 : Type (dropdown)
        tc = ws.cell(row=4, column=start_col)
        ws.merge_cells(start_row=4, start_column=start_col,
                       end_row=4, end_column=start_col + COLS_PM4 - 1)
        tc.fill = fill("F1F8E9")
        tc.alignment = align("center")
        dv_type.sqref = f"{dv_type.sqref} {col(start_col)}4" if dv_type.sqref else f"{col(start_col)}4"

        # Ligne 5 : Dom/Ext
        ec = ws.cell(row=5, column=start_col)
        ws.merge_cells(start_row=5, start_column=start_col,
                       end_row=5, end_column=start_col + COLS_PM4 - 1)
        ec.fill = fill("F9FBE7")
        ec.alignment = align("center")
        dv_domext.sqref = f"{dv_domext.sqref} {col(start_col)}5" if dv_domext.sqref else f"{col(start_col)}5"

        # Ligne 6 : Score
        sc6 = ws.cell(row=6, column=start_col)
        ws.merge_cells(start_row=6, start_column=start_col,
                       end_row=6, end_column=start_col + COLS_PM4 - 1)
        sc6.fill = fill("FFF9C4")
        sc6.alignment = align("center")

        # Ligne 7 : Sous-en-têtes
        sub_labels = ["Tit/Rem", "Min", "Buts", "PD"]
        sub_bgs    = ["A5D6A7", "C8E6C9", "DCEDC8", "F0F4C3"]
        for sc_idx in range(COLS_PM4):
            set_cell(ws, 7, start_col + sc_idx, sub_labels[sc_idx],
                     bg=sub_bgs[sc_idx], bold=True, halign="center",
                     size=8, valign="center")

        # Lignes joueurs
        for r in range(PLAYER_START_ROW, PLAYER_START_ROW + NB_PLAYERS):
            bg_row = "F9FBE7" if r % 2 == 0 else "FFFFFF"
            # Tit/Rem
            c_role = ws.cell(row=r, column=start_col)
            c_role.fill = fill(bg_row)
            c_role.alignment = align("center")
            # Minutes, Buts, PD
            for sc_idx in range(1, COLS_PM4):
                ccc = ws.cell(row=r, column=start_col + sc_idx)
                ccc.fill = fill(bg_row)
                ccc.alignment = align("center")

    # Dropdown rôle sur toute la plage titulaire/remplaçant
    role_ranges = []
    for m in range(NB_MATCHS_P4):
        sc = FIRST_MATCH_COL + m * COLS_PM4
        role_ranges.append(f"{col(sc)}{PLAYER_START_ROW}:{col(sc)}{PLAYER_START_ROW+NB_PLAYERS-1}")
    dv_role.sqref = " ".join(role_ranges)

    # Noms joueurs (colonne A, lignes 8+)
    for r in range(PLAYER_START_ROW, PLAYER_START_ROW + NB_PLAYERS):
        player_idx = r - PLAYER_START_ROW + 2
        set_cell(ws, r, 1,
                 formula=f"='Liste Joueurs'!B{player_idx}&\" \"&'Liste Joueurs'!C{player_idx}",
                 bg="CFD8DC", bold=True)

    ws.freeze_panes = f"B{PLAYER_START_ROW}"

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 5 : STATS GÉNÉRALES
# ─────────────────────────────────────────────────────────────────────────────

def build_page5(ws):
    ws.title = "Stats Générales"

    PLAYER_START_ROW = 4
    COLS_PM4 = 4
    SESSIONS_TOTAL = SESSIONS_PM + 2  # séances + TOT + %

    col_widths = [22, 12, 12, 12, 12, 14, 12, 10, 12, 16]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[col(i)].width = w

    headers = [
        "NOM / PRÉNOM",
        "Taux présence\nentraîn.",
        "Nb matchs\nconvoqué",
        "Taux\nconvocation",
        "Matchs\ntitulaire",
        "Matchs\nremplaçant",
        "Minutes\njouées",
        "Buts\nmarqués",
        "Passes\ndécisives",
        "Indice IRE\n(Conv/Présence)"
    ]

    # Titre
    merge(ws, 1,1,1,len(headers),
          "📊 STATISTIQUES GÉNÉRALES – Saison 2025/2026",
          bg="1A237E", bold=True, color="FFFFFF", size=13)
    ws.row_dimensions[1].height = 30

    # Légende IRE
    merge(ws, 2,1,2,len(headers),
          "Indice IRE :  🟢 1,00–1,10 = Justement convoqué   |   🟠 < 1,00 = Sous-convoqué   |   🔴 > 1,10 = Sur-convoqué",
          bg="E8EAF6", italic=True, color="3949AB")
    ws.row_dimensions[2].height = 20

    # En-têtes
    for i, h in enumerate(headers, 1):
        set_cell(ws, 3, i, h, bg="3949AB", bold=True, color="FFFFFF",
                 halign="center", valign="center", wrap=True, border=True)
    ws.row_dimensions[3].height = 50

    # Formules joueurs
    for r in range(PLAYER_START_ROW, PLAYER_START_ROW + NB_PLAYERS):
        player_idx = r - PLAYER_START_ROW + 2
        bg_row = "E8EAF6" if r % 2 == 0 else "FFFFFF"

        # Col A : Nom
        set_cell(ws, r, 1,
                 formula=f"='Liste Joueurs'!B{player_idx}&\" \"&'Liste Joueurs'!C{player_idx}",
                 bg="CFD8DC", bold=True)

        # Page 2 : joueurs en lignes 5 à 34
        p2_row = r - PLAYER_START_ROW + 5

        # Col B : Taux présence (moyenne des colonnes % de chaque mois)
        # Colonnes % : col 12, 23, 34, 45, 56, 67, 78, 89, 100, 111
        pct_cols = [2 + m * SESSIONS_TOTAL + SESSIONS_PM + 1 for m in range(len(MONTHS))]
        pct_refs = ",".join([f"'Présence Entraînement'!{col(c)}{p2_row}" for c in pct_cols])
        set_cell(ws, r, 2,
                 formula=f"=IFERROR(AVERAGEIF({{{pct_refs}}},\"<>\"),\"\")",
                 bg=bg_row, halign="center", num_format="0%")

        # Page 3 : joueurs en lignes 4 à 33
        p3_row = r - PLAYER_START_ROW + 4

        # Col C : Nb matchs convoqué
        set_cell(ws, r, 3,
                 formula=f"=COUNTIF('Convocations Matchs'!B{p3_row}:{col(NB_MATCHS_P3+1)}{p3_row},\"✅ Convoqué\")",
                 bg=bg_row, halign="center")

        # Col D : Taux convocation
        set_cell(ws, r, 4,
                 formula=f"=IFERROR(C{r}/COUNTA('Convocations Matchs'!B{p3_row}:{col(NB_MATCHS_P3+1)}{p3_row}),\"\")",
                 bg=bg_row, halign="center", num_format="0%")

        # Page 4 : joueurs en lignes 8 à 37
        p4_row = r - PLAYER_START_ROW + 8

        # Colonnes tit/rem/min/buts/pd pour chaque match
        tit_refs = ",".join([f"'Feuille de Match'!{col(2+m*COLS_PM4)}{p4_row}"   for m in range(NB_MATCHS_P4)])
        rem_refs = tit_refs  # même colonne, filtre différent
        min_refs = ",".join([f"'Feuille de Match'!{col(3+m*COLS_PM4)}{p4_row}"   for m in range(NB_MATCHS_P4)])
        but_refs = ",".join([f"'Feuille de Match'!{col(4+m*COLS_PM4)}{p4_row}"   for m in range(NB_MATCHS_P4)])
        pd_refs  = ",".join([f"'Feuille de Match'!{col(5+m*COLS_PM4)}{p4_row}"   for m in range(NB_MATCHS_P4)])

        # Col E : Titulaire
        set_cell(ws, r, 5, formula=f'=COUNTIF({{{tit_refs}}},"Titulaire")',
                 bg=bg_row, halign="center")

        # Col F : Remplaçant
        set_cell(ws, r, 6, formula=f'=COUNTIF({{{rem_refs}}},"Remplaçant")',
                 bg=bg_row, halign="center")

        # Col G : Minutes
        set_cell(ws, r, 7, formula=f"=SUM({{{min_refs}}})",
                 bg=bg_row, halign="center")

        # Col H : Buts
        set_cell(ws, r, 8, formula=f"=SUM({{{but_refs}}})",
                 bg=bg_row, halign="center")

        # Col I : Passes décisives
        set_cell(ws, r, 9, formula=f"=SUM({{{pd_refs}}})",
                 bg=bg_row, halign="center")

        # Col J : Indice IRE = D / B
        set_cell(ws, r, 10,
                 formula=f"=IFERROR(D{r}/B{r},\"\")",
                 bg=bg_row, halign="center", num_format="0.00", bold=True)

    # ── Mise en forme conditionnelle IRE ────────────────────────────────────
    ire_range = f"J{PLAYER_START_ROW}:J{PLAYER_START_ROW+NB_PLAYERS-1}"

    # Rouge : > 1,10
    ws.conditional_formatting.add(ire_range, FormulaRule(
        formula=[f"J{PLAYER_START_ROW}>1.1"],
        fill=fill("FFCDD2"),
        font=Font(color="B71C1C", bold=True)
    ))
    # Vert : entre 1 et 1,10
    ws.conditional_formatting.add(ire_range, FormulaRule(
        formula=[f"AND(J{PLAYER_START_ROW}>=1,J{PLAYER_START_ROW}<=1.1)"],
        fill=fill("C8E6C9"),
        font=Font(color="1B5E20", bold=True)
    ))
    # Orange : < 1
    ws.conditional_formatting.add(ire_range, FormulaRule(
        formula=[f"AND(J{PLAYER_START_ROW}<1,J{PLAYER_START_ROW}<>\"\")"],
        fill=fill("FFE0B2"),
        font=Font(color="E65100", bold=True)
    ))

    # Ligne moyenne équipe
    avg_row = PLAYER_START_ROW + NB_PLAYERS
    set_cell(ws, avg_row, 1, "MOYENNE ÉQUIPE",
             bg="37474F", bold=True, color="FFFFFF")
    avg_fmts = ["0%","0","0%","0","0","0","0","0","0.00"]
    for i, fmt in enumerate(avg_fmts, 2):
        cl = col(i)
        set_cell(ws, avg_row, i,
                 formula=f"=IFERROR(AVERAGE({cl}{PLAYER_START_ROW}:{cl}{PLAYER_START_ROW+NB_PLAYERS-1}),\"\")",
                 bg="263238", bold=True, color="FFFFFF",
                 halign="center", num_format=fmt)

    ws.freeze_panes = f"B{PLAYER_START_ROW}"

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    wb = Workbook()

    # Supprimer feuille par défaut
    default = wb.active
    wb.remove(default)

    # Créer les 5 feuilles
    ws1 = wb.create_sheet("Liste Joueurs")
    ws2 = wb.create_sheet("Présence Entraînement")
    ws3 = wb.create_sheet("Convocations Matchs")
    ws4 = wb.create_sheet("Feuille de Match")
    ws5 = wb.create_sheet("Stats Générales")

    print("📋 Page 1 : Liste des joueurs…")
    build_page1(ws1)

    print("📅 Page 2 : Présence entraînements…")
    build_page2(ws2)

    print("📣 Page 3 : Convocations matchs…")
    build_page3(ws3)

    print("⚽ Page 4 : Feuille de match…")
    build_page4(ws4)

    print("📊 Page 5 : Stats générales…")
    build_page5(ws5)

    output = "suivi_equipe.xlsx"
    wb.save(output)
    print(f"\n✅ Fichier créé : {output}")
    print("→ Double-cliquez sur le fichier pour l'ouvrir dans Numbers.")
    print("\nCONSEIL : Remplissez d'abord l'onglet 'Liste Joueurs'.")
    print("          Tous les autres onglets s'y réfèrent automatiquement.")

if __name__ == "__main__":
    main()
