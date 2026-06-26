/**
 * ╔══════════════════════════════════════════════════════════╗
 * ║   SUIVI D'ÉQUIPE – FOOTBALL JEUNES                       ║
 * ║   Entente Sportive de Genech                             ║
 * ╚══════════════════════════════════════════════════════════╝
 *
 * INSTRUCTIONS :
 * 1. Ouvrez un nouveau Google Sheets (sheets.new)
 * 2. Menu Extensions > Apps Script
 * 3. Supprimez tout le code existant
 * 4. Collez CE script entier
 * 5. Cliquez sur ▶ Exécuter  (fonction : createTemplate)
 * 6. Acceptez les autorisations Google
 */

// ─── CONFIG (à modifier selon votre équipe) ───────────────────────────────────
var CLUB      = "Entente Sportive de Genech";
var CATEGORIE = "U13";
var SAISON    = "2025/2026";
var NB_JOUEURS   = 25;
var SESSIONS_PM  = 9;    // séances par mois
var NB_MATCHS_P3 = 30;   // matchs convocations
var NB_MATCHS_P4 = 20;   // matchs feuille de match

// ─── COULEURS ─────────────────────────────────────────────────────────────────
var C_NAVY  = "#0D1B3E";
var C_HDR   = "#1A2E5A";
var C_BLUE  = "#1565C0";
var C_GOLD  = "#C9A84C";
var C_WHITE = "#FFFFFF";
var C_STRIP = "#E8EDF6";
var C_LIGHT = "#F0F4FA";
var C_NAME  = "#D6E4F0";
var C_TOT   = "#BBD6F5";
var C_FOOT  = "#263A5A";

var MONTH_COLORS = [
  "#1565C0","#6A1B9A","#00695C","#BF360C",
  "#4527A0","#AD1457","#2E7D32","#F57F17",
  "#0277BD","#558B2F"
];

var MONTHS = [
  "SEPTEMBRE","OCTOBRE","NOVEMBRE","DÉCEMBRE",
  "JANVIER","FÉVRIER","MARS","AVRIL","MAI","JUIN"
];

// ─── POINT D'ENTRÉE ───────────────────────────────────────────────────────────
function createTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var names = [
    "📋 Liste Joueurs",
    "📅 Présences Entraînement",
    "📣 Convocations Matchs",
    "⚽ Feuille de Match",
    "📊 Stats Générales"
  ];

  // Créer les feuilles manquantes
  names.forEach(function(n) {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });

  // Supprimer les feuilles non nommées
  ss.getSheets().forEach(function(s) {
    if (names.indexOf(s.getName()) === -1) ss.deleteSheet(s);
  });

  // Réordonner
  names.forEach(function(n, i) {
    ss.setActiveSheet(ss.getSheetByName(n));
    ss.moveActiveSheet(i + 1);
  });

  buildPage1(ss);
  buildPage2(ss);
  buildPage3(ss);
  buildPage4(ss);
  buildPage5(ss);

  SpreadsheetApp.flush();
  ss.setActiveSheet(ss.getSheetByName("📋 Liste Joueurs"));
  Browser.msgBox("✅ Modèle créé !\n\nCommencez par remplir « 📋 Liste Joueurs »\nTous les autres onglets s'y réfèrent automatiquement.");
}

// ═════════════════════════════════════════════════════════════════════════════
// PAGE 1 — LISTE DES JOUEURS
// ═════════════════════════════════════════════════════════════════════════════
function buildPage1(ss) {
  var sh = ss.getSheetByName("📋 Liste Joueurs");
  sh.clear();
  sh.clearConditionalFormatRules();

  // Largeurs
  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 50);
  sh.setColumnWidth(3, 170);
  sh.setColumnWidth(4, 170);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 200);
  sh.setColumnWidth(8, 30);

  // ── Bandeau titre ──────────────────────────────────────────────────────────
  sh.setRowHeight(1, 10);
  sh.setRowHeight(2, 60);
  sh.setRowHeight(3, 26);
  sh.setRowHeight(4, 10);
  sh.setRowHeight(5, 8);

  // Fond marine ligne 1-4
  sh.getRange(1,1,4,8).setBackground(C_NAVY);

  // Titre club
  var titre = sh.getRange(2,2,1,6).merge();
  setStyle(titre, CLUB + "  –  " + CATEGORIE, C_NAVY, C_GOLD, 16, true, "center");

  // Sous-titre
  var sousTitre = sh.getRange(3,2,1,6).merge();
  setStyle(sousTitre, "LISTE DES JOUEURS  –  Saison " + SAISON, C_NAVY, C_WHITE, 11, true, "center");

  // Trait or
  sh.getRange(5,1,1,8).setBackground(C_GOLD);

  // ── En-têtes ───────────────────────────────────────────────────────────────
  sh.setRowHeight(6, 28);
  var hdrs = ["", "N°", "Nom", "Prénom", "Date de naissance", "Poste", "Contact parent", ""];
  hdrs.forEach(function(h, i) {
    var c = sh.getRange(6, i+1);
    setStyle(c, h, C_HDR, C_WHITE, 10, true, "center");
    c.setBorder(true,true,true,true,null,null,"#C9A84C", SpreadsheetApp.BorderStyle.SOLID);
  });

  // ── Lignes joueurs ─────────────────────────────────────────────────────────
  for (var r = 7; r < 7 + NB_JOUEURS; r++) {
    sh.setRowHeight(r, 24);
    var bg = (r % 2 === 0) ? C_STRIP : C_WHITE;
    sh.getRange(r,1,1,8).setBackground(bg);
    setStyle(sh.getRange(r,2), r-6, bg, C_BLUE, 10, true, "center");
    sh.getRange(r,2,1,7).setBorder(true,true,true,true,true,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);
  }

  // ── Pied ───────────────────────────────────────────────────────────────────
  var foot = 7 + NB_JOUEURS;
  sh.setRowHeight(foot, 8);
  sh.getRange(foot,1,1,8).setBackground(C_GOLD);
  sh.setRowHeight(foot+1, 20);
  var pied = sh.getRange(foot+1,2,1,6).merge();
  setStyle(pied, "Saison " + SAISON + "  –  " + CLUB, C_NAVY, C_WHITE, 9, false, "center");
  pied.setFontStyle("italic");

  sh.setFrozenRows(6);
}

// ═════════════════════════════════════════════════════════════════════════════
// PAGE 2 — PRÉSENCES ENTRAÎNEMENT
// ═════════════════════════════════════════════════════════════════════════════
function buildPage2(ss) {
  var sh = ss.getSheetByName("📅 Présences Entraînement");
  sh.clear();
  sh.clearConditionalFormatRules();

  var COLS_PM      = SESSIONS_PM + 2;   // séances + TOT + %
  var FIRST_COL    = 2;
  var PLAYER_ROW   = 6;
  var LAST_PLAYER  = PLAYER_ROW + NB_JOUEURS - 1;
  var SEP_ROW      = LAST_PLAYER + 1;
  var TOT_NB_ROW   = SEP_ROW + 1;
  var TOT_PCT_ROW  = TOT_NB_ROW + 1;

  // Colonnes TOTAL ANNÉE
  var YEAR_NB_COL  = FIRST_COL + MONTHS.length * COLS_PM;
  var YEAR_PCT_COL = YEAR_NB_COL + 1;
  var TOTAL_COLS   = YEAR_PCT_COL;

  // ── Largeurs ───────────────────────────────────────────────────────────────
  sh.setColumnWidth(1, 170);
  for (var m = 0; m < MONTHS.length; m++) {
    var base = FIRST_COL + m * COLS_PM;
    for (var s = 0; s < SESSIONS_PM; s++) sh.setColumnWidth(base + s, 55);
    sh.setColumnWidth(base + SESSIONS_PM,     45);
    sh.setColumnWidth(base + SESSIONS_PM + 1, 45);
  }
  sh.setColumnWidth(YEAR_NB_COL,  70);
  sh.setColumnWidth(YEAR_PCT_COL, 70);

  // ── Titre ─────────────────────────────────────────────────────────────────
  sh.setRowHeight(1, 28);
  var titre = sh.getRange(1,1,1,TOTAL_COLS).merge();
  setStyle(titre, "📅  SUIVI DE PRÉSENCE AUX ENTRAÎNEMENTS  –  " + CLUB + "  –  " + CATEGORIE + "  –  Saison " + SAISON,
    C_NAVY, C_GOLD, 12, true, "center");

  // ── Légende ────────────────────────────────────────────────────────────────
  sh.setRowHeight(2, 22);
  var legends = [
    [1, 2, "P  =  Présent",            "#27AE60", C_WHITE],
    [3, 4, "AJ  =  Absence Justifiée", "#F1C40F", "#5D4E00"],
    [5, 6, "AI  =  Absence Injustifiée","#E74C3C", C_WHITE],
    [7, 8, "B  =  Blessé / Malade",    "#8E44AD", C_WHITE]
  ];
  legends.forEach(function(l) {
    var r = sh.getRange(2, l[0], 1, l[1]-l[0]+1).merge();
    setStyle(r, l[2], l[3], l[4], 9, true, "center");
  });

  // Trait or
  sh.setRowHeight(3, 6);
  sh.getRange(3,1,1,TOTAL_COLS).setBackground(C_GOLD);
  sh.setRowHeight(4, 10);

  // ── Ligne mois (ligne 4) ──────────────────────────────────────────────────
  sh.setRowHeight(4, 22);
  var colCursor = FIRST_COL;
  for (var m = 0; m < MONTHS.length; m++) {
    var mRange = sh.getRange(4, colCursor, 1, COLS_PM).merge();
    setStyle(mRange, MONTHS[m], MONTH_COLORS[m], C_WHITE, 10, true, "center");
    colCursor += COLS_PM;
  }
  var yearHdr = sh.getRange(4, YEAR_NB_COL, 1, 2).merge();
  setStyle(yearHdr, "TOTAL ANNÉE", C_GOLD, C_NAVY, 10, true, "center");

  // ── Sous-en-têtes séances (ligne 5) ───────────────────────────────────────
  sh.setRowHeight(5, 30);
  setStyle(sh.getRange(5,1), "NOM / PRÉNOM", C_HDR, C_WHITE, 10, true, "center");

  colCursor = FIRST_COL;
  for (var m = 0; m < MONTHS.length; m++) {
    for (var s = 1; s <= SESSIONS_PM; s++) {
      var c = sh.getRange(5, colCursor);
      setStyle(c, "Séance " + s, "#D6E8FF", C_NAVY, 8, true, "center");
      c.setBorder(true,true,true,true,null,null,"#AAAAAA", SpreadsheetApp.BorderStyle.SOLID_THIN);
      colCursor++;
    }
    var tot = sh.getRange(5, colCursor);
    setStyle(tot, "TOT", C_TOT, C_NAVY, 9, true, "center");
    colCursor++;
    var pct = sh.getRange(5, colCursor);
    setStyle(pct, "%", C_TOT, C_NAVY, 9, true, "center");
    colCursor++;
  }
  setStyle(sh.getRange(5, YEAR_NB_COL),  "Nb total", C_GOLD, C_NAVY, 8, true, "center");
  setStyle(sh.getRange(5, YEAR_PCT_COL), "% annuel", C_GOLD, C_NAVY, 8, true, "center");

  // ── Dropdown ──────────────────────────────────────────────────────────────
  var dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(["P","AJ","AI","B"], true)
    .setAllowInvalid(false).build();

  // ── Lignes joueurs ─────────────────────────────────────────────────────────
  for (var r = PLAYER_ROW; r <= LAST_PLAYER; r++) {
    sh.setRowHeight(r, 22);
    var pIdx = r - PLAYER_ROW + 7;  // ligne dans Liste Joueurs
    var bg = (r % 2 === 0) ? C_STRIP : C_WHITE;

    // Nom depuis Page 1
    var nomCell = sh.getRange(r, 1);
    nomCell.setFormula("='📋 Liste Joueurs'!C" + pIdx + "&\" \"&'📋 Liste Joueurs'!D" + pIdx);
    setStyle(nomCell, null, C_NAME, C_NAVY, 10, true, "left");

    var colC = FIRST_COL;
    var totRefs = [];

    for (var m = 0; m < MONTHS.length; m++) {
      var sesStart = colC;

      // Séances
      for (var s = 0; s < SESSIONS_PM; s++) {
        var cell = sh.getRange(r, colC);
        cell.setBackground(bg).setHorizontalAlignment("center");
        cell.setDataValidation(dv);
        cell.setBorder(true,true,true,true,null,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);
        colC++;
      }

      // TOT
      var sStart = colToLetter(sesStart) + r;
      var sEnd   = colToLetter(sesStart + SESSIONS_PM - 1) + r;
      var totCell = sh.getRange(r, colC);
      totCell.setFormula('=COUNTIF(' + sStart + ':' + sEnd + ',"P")');
      setStyle(totCell, null, C_TOT, C_NAVY, 9, true, "center");
      totRefs.push(colToLetter(colC) + r);
      colC++;

      // %
      var cntRange = colToLetter(sesStart) + r + ':' + colToLetter(sesStart + SESSIONS_PM - 1) + r;
      var pctCell = sh.getRange(r, colC);
      pctCell.setFormula('=IFERROR(' + colToLetter(colC-1) + r + '/COUNTA(' + cntRange + '),"")');
      setStyle(pctCell, null, C_TOT, C_NAVY, 9, true, "center");
      pctCell.setNumberFormat("0%");
      colC++;
    }

    // TOTAL ANNÉE – nb
    var nbAn = sh.getRange(r, YEAR_NB_COL);
    nbAn.setFormula("=" + totRefs.join("+"));
    setStyle(nbAn, null, "#FFF5CC", C_NAVY, 10, true, "center");
    nbAn.setFontWeight("bold");

    // TOTAL ANNÉE – %
    var allSes = [];
    var c3 = FIRST_COL;
    for (var m = 0; m < MONTHS.length; m++) {
      allSes.push(colToLetter(c3) + r + ':' + colToLetter(c3 + SESSIONS_PM - 1) + r);
      c3 += COLS_PM;
    }
    var pctAn = sh.getRange(r, YEAR_PCT_COL);
    pctAn.setFormula('=IFERROR(' + colToLetter(YEAR_NB_COL) + r + '/COUNTA(' + allSes.join(',') + '),"")');
    setStyle(pctAn, null, "#FFF5CC", C_NAVY, 10, true, "center");
    pctAn.setNumberFormat("0%");
    pctAn.setFontWeight("bold");
  }

  // ── Ligne séparatrice ──────────────────────────────────────────────────────
  sh.setRowHeight(SEP_ROW, 6);
  sh.getRange(SEP_ROW, 1, 1, TOTAL_COLS).setBackground("#C9D8EE");

  // ── TOTAUX BAS : Nb présents par séance ───────────────────────────────────
  sh.setRowHeight(TOT_NB_ROW,  22);
  sh.setRowHeight(TOT_PCT_ROW, 22);
  setStyle(sh.getRange(TOT_NB_ROW,  1), "Nb présents / séance",  C_FOOT, C_WHITE, 9, true, "center");
  setStyle(sh.getRange(TOT_PCT_ROW, 1), "% présence / séance",   C_FOOT, C_GOLD,  9, true, "center");

  var colC2 = FIRST_COL;
  for (var m = 0; m < MONTHS.length; m++) {
    for (var s = 0; s < SESSIONS_PM; s++) {
      var cLtr = colToLetter(colC2);
      var nbCell = sh.getRange(TOT_NB_ROW, colC2);
      nbCell.setFormula('=COUNTIF(' + cLtr + PLAYER_ROW + ':' + cLtr + LAST_PLAYER + ',"P")');
      setStyle(nbCell, null, C_FOOT, C_WHITE, 9, true, "center");
      nbCell.setBorder(true,true,true,true,null,null,C_NAVY, SpreadsheetApp.BorderStyle.SOLID_THIN);

      var pCell = sh.getRange(TOT_PCT_ROW, colC2);
      pCell.setFormula('=IFERROR(' + cLtr + TOT_NB_ROW + '/COUNTA(' + cLtr + PLAYER_ROW + ':' + cLtr + LAST_PLAYER + '),"")');
      setStyle(pCell, null, C_FOOT, C_GOLD, 9, true, "center");
      pCell.setNumberFormat("0%");
      pCell.setBorder(true,true,true,true,null,null,C_NAVY, SpreadsheetApp.BorderStyle.SOLID_THIN);
      colC2++;
    }
    // TOT + % colonnes
    sh.getRange(TOT_NB_ROW,  colC2).setBackground(C_FOOT);
    sh.getRange(TOT_PCT_ROW, colC2).setBackground(C_FOOT);
    colC2++;
    sh.getRange(TOT_NB_ROW,  colC2).setBackground(C_FOOT);
    sh.getRange(TOT_PCT_ROW, colC2).setBackground(C_FOOT);
    colC2++;
  }

  // TOTAL ANNÉE bas
  var ynbL = colToLetter(YEAR_NB_COL);
  var nb = sh.getRange(TOT_NB_ROW, YEAR_NB_COL);
  nb.setFormula('=SUM(' + ynbL + PLAYER_ROW + ':' + ynbL + LAST_PLAYER + ')');
  setStyle(nb, null, C_GOLD, C_NAVY, 10, true, "center");
  nb.setFontWeight("bold");

  var ypctL = colToLetter(YEAR_PCT_COL);
  var pctAnn = sh.getRange(TOT_PCT_ROW, YEAR_PCT_COL);
  pctAnn.setFormula('=IFERROR(AVERAGE(' + ypctL + PLAYER_ROW + ':' + ypctL + LAST_PLAYER + '),"")');
  setStyle(pctAnn, null, C_GOLD, C_NAVY, 10, true, "center");
  pctAnn.setNumberFormat("0%");
  pctAnn.setFontWeight("bold");

  // ── Mise en couleur conditionnelle (P/AJ/AI/B) ────────────────────────────
  var sessionRanges = [];
  var c4 = FIRST_COL;
  for (var m = 0; m < MONTHS.length; m++) {
    sessionRanges.push(sh.getRange(PLAYER_ROW, c4, NB_JOUEURS, SESSIONS_PM));
    c4 += COLS_PM;
  }

  var cfRules = [];
  var cfValues = [
    ["P",  "#27AE60", C_WHITE],
    ["AJ", "#F1C40F", "#5D4E00"],
    ["AI", "#E74C3C", C_WHITE],
    ["B",  "#8E44AD", C_WHITE]
  ];
  cfValues.forEach(function(v) {
    sessionRanges.forEach(function(rng) {
      cfRules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(v[0])
        .setBackground(v[1])
        .setFontColor(v[2])
        .setBold(true)
        .setRanges([rng])
        .build());
    });
  });
  sh.setConditionalFormatRules(cfRules);

  sh.setFrozenRows(5);
  sh.setFrozenColumns(1);
}

// ═════════════════════════════════════════════════════════════════════════════
// PAGE 3 — CONVOCATIONS MATCHS
// ═════════════════════════════════════════════════════════════════════════════
function buildPage3(ss) {
  var sh = ss.getSheetByName("📣 Convocations Matchs");
  sh.clear();
  sh.clearConditionalFormatRules();

  var PLAYER_ROW  = 6;
  var LAST_PLAYER = PLAYER_ROW + NB_JOUEURS - 1;
  var SEP_ROW     = LAST_PLAYER + 1;
  var TOT_NB_ROW  = SEP_ROW + 1;
  var TOT_PCT_ROW = TOT_NB_ROW + 1;
  var TOTAL_COLS  = NB_MATCHS_P3 + 2;

  sh.setColumnWidth(1, 170);
  for (var m = 1; m <= NB_MATCHS_P3; m++) sh.setColumnWidth(m+1, 90);

  // Titre
  sh.setRowHeight(1, 28);
  setStyle(sh.getRange(1,1,1,TOTAL_COLS).merge(),
    "📣  CONVOCATIONS MATCHS  –  " + CLUB + "  –  " + CATEGORIE + "  –  Saison " + SAISON,
    C_NAVY, C_GOLD, 12, true, "center");

  // Légende
  sh.setRowHeight(2, 22);
  var leg = [
    [1,2,"✅  Convoqué","#27AE60",C_WHITE],
    [3,4,"❌  Non convoqué","#E74C3C",C_WHITE],
    [5,7,"⚠️  Convoqué absent","#F39C12",C_WHITE]
  ];
  leg.forEach(function(l) {
    setStyle(sh.getRange(2,l[0],1,l[1]-l[0]+1).merge(), l[2], l[3], l[4], 9, true, "center");
  });

  sh.setRowHeight(3, 6); sh.getRange(3,1,1,TOTAL_COLS).setBackground(C_GOLD);
  sh.setRowHeight(4, 10);
  sh.setRowHeight(5, 26);

  // En-tête noms
  setStyle(sh.getRange(5,1), "NOM / PRÉNOM", C_HDR, C_WHITE, 10, true, "center");
  sh.getRange(5,1).setBorder(true,true,true,true,null,null,C_GOLD, SpreadsheetApp.BorderStyle.SOLID);

  var mColors = ["#1A5276","#1B4F72","#154360","#0E6655","#145A32",
                 "#4A235A","#6E2F7A","#7B241C","#78281F","#1F618D"];
  for (var m = 1; m <= NB_MATCHS_P3; m++) {
    var hdr = sh.getRange(5, m+1);
    setStyle(hdr, "Match " + m, mColors[(m-1)%mColors.length], C_WHITE, 9, true, "center");
    hdr.setBorder(true,true,true,true,null,null,"#AAAAAA", SpreadsheetApp.BorderStyle.SOLID_THIN);
  }

  // Dropdown
  var dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(["✅ Convoqué","❌ Non convoqué","⚠️ Convoqué absent"], true)
    .setAllowInvalid(false).build();

  // Lignes joueurs
  for (var r = PLAYER_ROW; r <= LAST_PLAYER; r++) {
    sh.setRowHeight(r, 22);
    var pIdx = r - PLAYER_ROW + 7;
    var bg = (r % 2 === 0) ? C_STRIP : C_WHITE;
    var nomCell = sh.getRange(r, 1);
    nomCell.setFormula("='📋 Liste Joueurs'!C" + pIdx + "&\" \"&'📋 Liste Joueurs'!D" + pIdx);
    setStyle(nomCell, null, C_NAME, C_NAVY, 10, true, "left");
    nomCell.setFontWeight("bold");
    for (var m = 1; m <= NB_MATCHS_P3; m++) {
      var c = sh.getRange(r, m+1);
      c.setBackground(bg).setHorizontalAlignment("center");
      c.setDataValidation(dv);
      c.setBorder(true,true,true,true,null,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);
    }
  }

  // Totaux bas
  sh.setRowHeight(SEP_ROW, 6); sh.getRange(SEP_ROW,1,1,TOTAL_COLS).setBackground("#C9D8EE");
  sh.setRowHeight(TOT_NB_ROW, 22); sh.setRowHeight(TOT_PCT_ROW, 22);
  setStyle(sh.getRange(TOT_NB_ROW, 1),  "Nb convoqués / match",  C_FOOT, C_WHITE, 9, true, "center");
  setStyle(sh.getRange(TOT_PCT_ROW, 1), "% convocation / match", C_FOOT, C_GOLD,  9, true, "center");

  for (var m = 1; m <= NB_MATCHS_P3; m++) {
    var cL = colToLetter(m+1);
    var nbC = sh.getRange(TOT_NB_ROW, m+1);
    nbC.setFormula('=COUNTIF(' + cL + PLAYER_ROW + ':' + cL + LAST_PLAYER + ',"✅ Convoqué")');
    setStyle(nbC, null, C_FOOT, C_WHITE, 9, true, "center");
    nbC.setBorder(true,true,true,true,null,null,C_NAVY, SpreadsheetApp.BorderStyle.SOLID_THIN);

    var pC = sh.getRange(TOT_PCT_ROW, m+1);
    pC.setFormula('=IFERROR(' + cL + TOT_NB_ROW + '/COUNTA(' + cL + PLAYER_ROW + ':' + cL + LAST_PLAYER + '),"")');
    setStyle(pC, null, C_FOOT, C_GOLD, 9, true, "center");
    pC.setNumberFormat("0%");
    pC.setBorder(true,true,true,true,null,null,C_NAVY, SpreadsheetApp.BorderStyle.SOLID_THIN);
  }

  // Couleur conditionnelle convocations
  var convRange = sh.getRange(PLAYER_ROW, 2, NB_JOUEURS, NB_MATCHS_P3);
  var cfConv = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("✅ Convoqué").setBackground("#D5F5E3").setFontColor("#1E8449").setBold(true).setRanges([convRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("❌ Non convoqué").setBackground("#FADBD8").setFontColor("#C0392B").setBold(true).setRanges([convRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("⚠️ Convoqué absent").setBackground("#FEF9E7").setFontColor("#D68910").setBold(true).setRanges([convRange]).build()
  ];
  sh.setConditionalFormatRules(cfConv);

  sh.setFrozenRows(5);
  sh.setFrozenColumns(1);
}

// ═════════════════════════════════════════════════════════════════════════════
// PAGE 4 — FEUILLE DE MATCH
// ═════════════════════════════════════════════════════════════════════════════
function buildPage4(ss) {
  var sh = ss.getSheetByName("⚽ Feuille de Match");
  sh.clear();
  sh.clearConditionalFormatRules();

  var COLS_PM4   = 4;
  var PLAYER_ROW = 9;
  var FIRST_COL  = 2;
  var TOTAL_COLS = 1 + NB_MATCHS_P4 * COLS_PM4;

  sh.setColumnWidth(1, 170);

  // Titre
  sh.setRowHeight(1, 28);
  setStyle(sh.getRange(1,1,1,TOTAL_COLS).merge(),
    "⚽  FEUILLE DE MATCH  –  " + CLUB + "  –  " + CATEGORIE + "  –  Saison " + SAISON,
    "#1B4F1F", C_GOLD, 12, true, "center");
  sh.setRowHeight(2, 6); sh.getRange(2,1,1,TOTAL_COLS).setBackground(C_GOLD);

  // Labels colonne A
  var aLabels = {3:"MATCH", 4:"Date", 5:"Type", 6:"Dom / Ext", 7:"Score", 8:"NOM / PRÉNOM"};
  var aHts    = {3:22, 4:22, 5:22, 6:22, 7:22, 8:26};
  for (var row in aLabels) {
    sh.setRowHeight(parseInt(row), aHts[row]);
    setStyle(sh.getRange(parseInt(row), 1), aLabels[row], C_HDR, C_WHITE, 9, true, "center");
  }

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Championnat","Amical","Coupe","Tournoi"], true)
    .setAllowInvalid(false).build();
  var domExtRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Domicile","Extérieur","Terrain neutre"], true)
    .setAllowInvalid(false).build();
  var roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Titulaire","Remplaçant","N/A"], true)
    .setAllowInvalid(false).build();

  var mHdrColors = ["#1A5276","#1B4F72","#154360","#0E6655","#145A32",
                    "#4A235A","#6E2F7A","#7B241C","#78281F","#1F618D",
                    "#17202A","#1B2631","#0B5345","#4D5656","#4A4A4A",
                    "#2C3E50","#641E16","#196F3D","#7D6608","#0D47A1"];

  for (var m = 0; m < NB_MATCHS_P4; m++) {
    var startCol = FIRST_COL + m * COLS_PM4;
    var mc = mHdrColors[m % mHdrColors.length];

    // Largeurs
    sh.setColumnWidth(startCol,   90);
    sh.setColumnWidth(startCol+1, 60);
    sh.setColumnWidth(startCol+2, 55);
    sh.setColumnWidth(startCol+3, 65);

    // Ligne 3 : Match N
    setStyle(sh.getRange(3, startCol, 1, COLS_PM4).merge(), "Match " + (m+1), mc, C_WHITE, 10, true, "center");

    // Ligne 4 : Date
    var dateR = sh.getRange(4, startCol, 1, COLS_PM4).merge();
    dateR.setBackground("#EBF5FB").setHorizontalAlignment("center").setNumberFormat("DD/MM/YYYY");

    // Ligne 5 : Type
    var typeR = sh.getRange(5, startCol, 1, COLS_PM4).merge();
    typeR.setBackground("#E8F8F5").setHorizontalAlignment("center");
    typeR.setDataValidation(typeRule);

    // Ligne 6 : Dom/Ext
    var domR = sh.getRange(6, startCol, 1, COLS_PM4).merge();
    domR.setBackground("#FEF9E7").setHorizontalAlignment("center");
    domR.setDataValidation(domExtRule);

    // Ligne 7 : Score
    var scoreR = sh.getRange(7, startCol, 1, COLS_PM4).merge();
    scoreR.setBackground("#FDEDEC").setHorizontalAlignment("center");
    scoreR.setFontWeight("bold").setFontColor("#C0392B").setFontSize(12);

    // Ligne 8 : Sous-en-têtes
    var subLabels = ["Tit / Rem","Min","Buts","Passes"];
    var subBgs    = ["#A9DFBF","#A9CCE3","#F9E79F","#D2B4DE"];
    for (var s = 0; s < COLS_PM4; s++) {
      var subC = sh.getRange(8, startCol+s);
      setStyle(subC, subLabels[s], subBgs[s], C_NAVY, 8, true, "center");
      subC.setBorder(true,true,true,true,null,null,"#AAAAAA", SpreadsheetApp.BorderStyle.SOLID_THIN);
    }

    // Lignes joueurs
    for (var r = PLAYER_ROW; r < PLAYER_ROW + NB_JOUEURS; r++) {
      var bg = (r % 2 === 0) ? C_STRIP : C_WHITE;
      for (var s = 0; s < COLS_PM4; s++) {
        var c = sh.getRange(r, startCol + s);
        c.setBackground(bg).setHorizontalAlignment("center");
        c.setBorder(true,true,true,true,null,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);
      }
      sh.getRange(r, startCol).setDataValidation(roleRule);
    }
  }

  // Noms joueurs
  for (var r = PLAYER_ROW; r < PLAYER_ROW + NB_JOUEURS; r++) {
    sh.setRowHeight(r, 22);
    var pIdx = r - PLAYER_ROW + 7;
    var nomCell = sh.getRange(r, 1);
    nomCell.setFormula("='📋 Liste Joueurs'!C" + pIdx + "&\" \"&'📋 Liste Joueurs'!D" + pIdx);
    setStyle(nomCell, null, C_NAME, C_NAVY, 10, true, "left");
    nomCell.setFontWeight("bold");
  }

  sh.setFrozenRows(8);
  sh.setFrozenColumns(1);
}

// ═════════════════════════════════════════════════════════════════════════════
// PAGE 5 — STATS GÉNÉRALES
// ═════════════════════════════════════════════════════════════════════════════
function buildPage5(ss) {
  var sh = ss.getSheetByName("📊 Stats Générales");
  sh.clear();
  sh.clearConditionalFormatRules();

  var PLAYER_ROW  = 6;
  var LAST_PLAYER = PLAYER_ROW + NB_JOUEURS - 1;
  var COLS_PM4    = 4;
  var COLS_PM_P2  = SESSIONS_PM + 2;
  var TOTAL_COLS  = 10;

  var colWidths = [170,100,100,100,100,110,90,80,100,120];
  colWidths.forEach(function(w, i) { sh.setColumnWidth(i+1, w); });

  var headers = [
    "NOM / PRÉNOM","Taux présence\nentraîn.","Matchs\nconvoqué",
    "Taux\nconvocation","Matchs\ntitulaire","Matchs\nremplaçant",
    "Minutes\njouées","Buts","Passes\ndécisives","Indice IRE\n(Conv ÷ Présence)"
  ];

  // Titre
  sh.setRowHeight(1, 28);
  setStyle(sh.getRange(1,1,1,TOTAL_COLS).merge(),
    "📊  STATISTIQUES GÉNÉRALES  –  " + CLUB + "  –  " + CATEGORIE + "  –  Saison " + SAISON,
    C_NAVY, C_GOLD, 12, true, "center");

  // Explication IRE
  sh.setRowHeight(2, 20);
  setStyle(sh.getRange(2,1,1,TOTAL_COLS).merge(),
    "Indice IRE :   🟢 1,00 – 1,10 = Justement convoqué     🟠 < 1,00 = Sous-convoqué     🔴 > 1,10 = Sur-convoqué",
    "#EEF2FA", C_HDR, 9, false, "center");
  sh.getRange(2,1).setFontStyle("italic");

  sh.setRowHeight(3, 6); sh.getRange(3,1,1,TOTAL_COLS).setBackground(C_GOLD);
  sh.setRowHeight(4, 10);

  // En-têtes
  sh.setRowHeight(5, 50);
  headers.forEach(function(h, i) {
    var c = sh.getRange(5, i+1);
    setStyle(c, h, C_HDR, C_WHITE, 9, true, "center");
    c.setWrap(true).setVerticalAlignment("middle");
    c.setBorder(true,true,true,true,null,null,C_GOLD, SpreadsheetApp.BorderStyle.SOLID);
  });

  // Lignes joueurs
  for (var r = PLAYER_ROW; r <= LAST_PLAYER; r++) {
    sh.setRowHeight(r, 24);
    var pIdx = r - PLAYER_ROW + 7;
    var bg = (r % 2 === 0) ? C_STRIP : C_WHITE;

    // Nom
    var nomCell = sh.getRange(r, 1);
    nomCell.setFormula("='📋 Liste Joueurs'!C" + pIdx + "&\" \"&'📋 Liste Joueurs'!D" + pIdx);
    setStyle(nomCell, null, C_NAME, C_NAVY, 10, true, "left");
    nomCell.setFontWeight("bold");

    // Page 2 – colonnes % de chaque mois (col FIRST_COL + m*COLS_PM + SESSIONS_PM + 1)
    var pctRefs = [];
    for (var m = 0; m < MONTHS.length; m++) {
      pctRefs.push("'📅 Présences Entraînement'!" + colToLetter(2 + m * COLS_PM_P2 + SESSIONS_PM + 1) + r);
    }
    var b = sh.getRange(r, 2);
    b.setFormula("=IFERROR(AVERAGEIF({" + pctRefs.join(",") + "},\"<>\"),\"\")");
    setStyle(b, null, bg, C_NAVY, 10, false, "center");
    b.setNumberFormat("0%");
    b.setBorder(true,true,true,true,null,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);

    // Page 3 – même numéro de ligne
    var c3 = sh.getRange(r, 3);
    c3.setFormula("=COUNTIF('📣 Convocations Matchs'!B" + r + ":" + colToLetter(NB_MATCHS_P3+1) + r + ",\"✅ Convoqué\")");
    setStyle(c3, null, bg, C_NAVY, 10, false, "center");
    c3.setBorder(true,true,true,true,null,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);

    var d = sh.getRange(r, 4);
    d.setFormula("=IFERROR(C" + r + "/COUNTA('📣 Convocations Matchs'!B" + r + ":" + colToLetter(NB_MATCHS_P3+1) + r + "),\"\")");
    setStyle(d, null, bg, C_NAVY, 10, false, "center");
    d.setNumberFormat("0%");
    d.setBorder(true,true,true,true,null,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);

    // Page 4 – ligne joueur décalée de +3 (PLAYER_ROW_P4=9, PLAYER_ROW_P5=6)
    var p4r = r + 3;
    var tit=[], min_=[], buts=[], pd=[];
    for (var m = 0; m < NB_MATCHS_P4; m++) {
      var sc = 2 + m * COLS_PM4;
      tit.push( "'⚽ Feuille de Match'!" + colToLetter(sc)   + p4r);
      min_.push("'⚽ Feuille de Match'!" + colToLetter(sc+1) + p4r);
      buts.push("'⚽ Feuille de Match'!" + colToLetter(sc+2) + p4r);
      pd.push(  "'⚽ Feuille de Match'!" + colToLetter(sc+3) + p4r);
    }

    var cols56789 = [
      ['=COUNTIF({' + tit.join(",")  + '},"Titulaire")',   bg, "0"],
      ['=COUNTIF({' + tit.join(",")  + '},"Remplaçant")',  bg, "0"],
      ['=SUM({'     + min_.join(",") + '})',                bg, "0"],
      ['=SUM({'     + buts.join(",") + '})',                bg, "0"],
      ['=SUM({'     + pd.join(",")   + '})',                bg, "0"]
    ];
    cols56789.forEach(function(cfg, i) {
      var cell = sh.getRange(r, 5+i);
      cell.setFormula(cfg[0]);
      setStyle(cell, null, cfg[1], C_NAVY, 10, false, "center");
      cell.setNumberFormat(cfg[2]);
      cell.setBorder(true,true,true,true,null,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);
    });

    // IRE
    var ire = sh.getRange(r, 10);
    ire.setFormula("=IFERROR(D" + r + "/B" + r + ",\"\")");
    setStyle(ire, null, bg, C_NAVY, 10, true, "center");
    ire.setNumberFormat("0.00");
    ire.setBorder(true,true,true,true,null,null,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);
  }

  // ── Formatage conditionnel IRE ─────────────────────────────────────────────
  var ireRange = sh.getRange(PLAYER_ROW, 10, NB_JOUEURS, 1);
  var cfRules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1.10)
      .setBackground("#FFCDD2").setFontColor("#C0392B").setBold(true).setRanges([ireRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1, 1.10)
      .setBackground("#D5F5E3").setFontColor("#1E8449").setBold(true).setRanges([ireRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1)
      .setBackground("#FAD7A0").setFontColor("#D35400").setBold(true).setRanges([ireRange]).build()
  ];
  sh.setConditionalFormatRules(cfRules);

  // Ligne moyenne
  var avgRow = LAST_PLAYER + 2;
  sh.setRowHeight(LAST_PLAYER+1, 6); sh.getRange(LAST_PLAYER+1,1,1,TOTAL_COLS).setBackground("#C9D8EE");
  sh.setRowHeight(avgRow, 26);
  setStyle(sh.getRange(avgRow, 1), "MOYENNE ÉQUIPE", C_FOOT, C_GOLD, 10, true, "center");

  var avgFmts = ["0%","0","0%","0","0","0","0","0","0.00"];
  for (var i = 0; i < avgFmts.length; i++) {
    var cl = colToLetter(i+2);
    var avgC = sh.getRange(avgRow, i+2);
    avgC.setFormula("=IFERROR(AVERAGE(" + cl + PLAYER_ROW + ":" + cl + LAST_PLAYER + "),\"\")");
    setStyle(avgC, null, C_FOOT, C_WHITE, 10, true, "center");
    avgC.setNumberFormat(avgFmts[i]);
    avgC.setBorder(true,true,true,true,null,null,C_NAVY, SpreadsheetApp.BorderStyle.SOLID_THIN);
  }

  sh.setFrozenRows(5);
  sh.setFrozenColumns(1);
}

// ═════════════════════════════════════════════════════════════════════════════
// UTILITAIRES
// ═════════════════════════════════════════════════════════════════════════════

function setStyle(range, value, bg, fg, size, bold, halign) {
  if (value !== null && value !== undefined) range.setValue(value);
  if (bg) range.setBackground(bg);
  range.setFontColor(fg || "#000000");
  range.setFontSize(size || 10);
  range.setFontWeight(bold ? "bold" : "normal");
  range.setHorizontalAlignment(halign || "left");
  range.setVerticalAlignment("middle");
  range.setFontFamily("Arial");
  return range;
}

function colToLetter(n) {
  var s = "";
  while (n > 0) {
    var mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
