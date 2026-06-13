# -*- coding: utf-8 -*-
"""Genere un fichier Excel (.xlsx) de suivi equipements + vacances."""
import zipfile
from xml.sax.saxutils import escape

PLAYERS = [
    "Tim", "Romain", "Jugboom", "Mathéo", "Clément", "Coco", "Capi", "Vaev",
    "Arthur", "Rémi", "Gabyn", "Hugo H", "Mathys", "Jules", "Gus", "Gio",
    "Dany", "Cyp", "Gaspard", "JP", "Thomas", "Julien", "Elio",
]

# Colonnes semaines de vacances (lundi -> dimanche) couvrant juillet/aout 2026
WEEKS = [
    "29 juin\n05 juil", "06\n12 juil", "13\n19 juil", "20\n26 juil",
    "27 juil\n02 août", "03\n09 août", "10\n16 août", "17\n23 août",
    "24\n30 août",
]

# En-tetes fixes
FIXED_HEADERS = ["Joueur", "Taille Haut", "Taille Bas", "N° Équipement"]
HEADERS = FIXED_HEADERS + WEEKS

NCOLS = len(HEADERS)


def col_letter(idx):  # 0-based -> A, B, ...
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def cell(ref, style, text=None):
    if text is None or text == "":
        return f'<c r="{ref}" s="{style}"/>'
    txt = escape(str(text))
    return (f'<c r="{ref}" s="{style}" t="inlineStr">'
            f'<is><t xml:space="preserve">{txt}</t></is></c>')


# ---------------------------------------------------------------- worksheet
rows_xml = []

# Ligne 1 : titre (fusionne sur toutes les colonnes)
r = 1
title_cells = [cell(f"A{r}", 6, "SUIVI ÉQUIPEMENTS & VACANCES — JUILLET / AOÛT 2026")]
for c in range(1, NCOLS):
    title_cells.append(f'<c r="{col_letter(c)}{r}" s="6"/>')
rows_xml.append(f'<row r="{r}" ht="30" customHeight="1">{"".join(title_cells)}</row>')

# Ligne 2 : en-tetes
r = 2
hdr_cells = []
for c, h in enumerate(HEADERS):
    style = 2 if c < len(FIXED_HEADERS) else 3  # bleu fonce / bleu clair
    text = h.replace("\n", "&#10;")  # garder le saut de ligne dans inlineStr
    ref = f"{col_letter(c)}{r}"
    # construire manuellement pour conserver le saut de ligne
    hdr_cells.append(
        f'<c r="{ref}" s="{style}" t="inlineStr"><is>'
        f'<t xml:space="preserve">{escape(h)}</t></is></c>')
rows_xml.append(f'<row r="{r}" ht="34" customHeight="1">{"".join(hdr_cells)}</row>')

# Lignes joueurs
for i, name in enumerate(PLAYERS):
    r = 3 + i
    cells = [cell(f"A{r}", 4, name)]  # nom : gras a gauche
    for c in range(1, NCOLS):
        cells.append(cell(f"{col_letter(c)}{r}", 5))  # cellules vides bordees
    rows_xml.append(f'<row r="{r}" ht="22" customHeight="1">{"".join(cells)}</row>')

last_row = 2 + len(PLAYERS)
last_col = col_letter(NCOLS - 1)
dim = f"A1:{last_col}{last_row}"

# Largeurs de colonnes
cols = ['<col min="1" max="1" width="14" customWidth="1"/>',
        '<col min="2" max="3" width="11" customWidth="1"/>',
        '<col min="4" max="4" width="13" customWidth="1"/>',
        f'<col min="5" max="{NCOLS}" width="10" customWidth="1"/>']

merge = f'<mergeCells count="1"><mergeCell ref="A1:{last_col}1"/></mergeCells>'

sheet = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="{dim}"/>
<sheetViews><sheetView workbookViewId="0">
<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>
</sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>{''.join(cols)}</cols>
<sheetData>{''.join(rows_xml)}</sheetData>
{merge}
</worksheet>'''

# ---------------------------------------------------------------- styles
styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border>
<left style="thin"><color rgb="FFB0B0B0"/></left>
<right style="thin"><color rgb="FFB0B0B0"/></right>
<top style="thin"><color rgb="FFB0B0B0"/></top>
<bottom style="thin"><color rgb="FFB0B0B0"/></bottom>
</border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Suivi équipes" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

out = "Suivi_Equipements_Vacances.xlsx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", root_rels)
    z.writestr("xl/workbook.xml", workbook)
    z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
    z.writestr("xl/styles.xml", styles)
    z.writestr("xl/worksheets/sheet1.xml", sheet)

print("Cree:", out, "—", len(PLAYERS), "joueurs,", NCOLS, "colonnes")
