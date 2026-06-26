/**
 * MODÈLE DE SUIVI D'ÉQUIPE - FOOTBALL JEUNES
 * Copiez ce script dans Outils > Éditeur de scripts, puis lancez createTemplate()
 */

function createTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Supprimer les feuilles existantes sauf la première
  var sheets = ss.getSheets();
  var sheetNames = ['Liste Joueurs', 'Présence Entraînement', 'Convocations Matchs', 'Feuille de Match', 'Stats Générales'];

  // Créer toutes les feuilles nécessaires
  sheetNames.forEach(function(name) {
    var existing = ss.getSheetByName(name);
    if (!existing) {
      ss.insertSheet(name);
    }
  });

  // Supprimer les feuilles non nommées (Sheet1, etc.)
  ss.getSheets().forEach(function(s) {
    if (sheetNames.indexOf(s.getName()) === -1) {
      ss.deleteSheet(s);
    }
  });

  // Réordonner les feuilles
  sheetNames.forEach(function(name, i) {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(i + 1);
  });

  buildPage1(ss);
  buildPage2(ss);
  buildPage3(ss);
  buildPage4(ss);
  buildPage5(ss);

  SpreadsheetApp.flush();
  Browser.msgBox('✅ Modèle créé avec succès ! Remplissez d\'abord la liste des joueurs (Page 1), puis utilisez les autres onglets.');
}

// ─────────────────────────────────────────────
// PAGE 1 : LISTE DES JOUEURS
// ─────────────────────────────────────────────
function buildPage1(ss) {
  var sh = ss.getSheetByName('Liste Joueurs');
  sh.clear();
  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 120);

  // Titre
  sh.getRange('A1:F1').merge().setValue('📋 LISTE DES JOUEURS').setBackground('#1a237e').setFontColor('#ffffff').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

  // Sous-titre équipe
  sh.getRange('A2:F2').merge().setValue('Saison 2025/2026 – FC [Nom du club] – Catégorie [U13/U15...]').setBackground('#e8eaf6').setFontSize(11).setHorizontalAlignment('center').setFontStyle('italic');

  // En-têtes
  var headers = ['N°', 'Nom', 'Prénom', 'Date de naissance', 'Poste', 'Contact parent'];
  var hRange = sh.getRange(3, 1, 1, headers.length);
  hRange.setValues([headers]).setBackground('#3949ab').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 30 lignes joueurs alternées
  for (var r = 4; r <= 33; r++) {
    sh.getRange(r, 1).setValue(r - 3);
    var bg = (r % 2 === 0) ? '#e8eaf6' : '#ffffff';
    sh.getRange(r, 1, 1, 6).setBackground(bg);
  }

  sh.setFrozenRows(3);
  sh.getRange('A3:F33').setBorder(true, true, true, true, true, true);

  // Bordure extérieure plus épaisse
  sh.getRange('A3:F33').setBorder(true, true, true, true, null, null, '#1a237e', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// ─────────────────────────────────────────────
// PAGE 2 : PRÉSENCE ENTRAÎNEMENTS
// ─────────────────────────────────────────────
function buildPage2(ss) {
  var sh = ss.getSheetByName('Présence Entraînement');
  sh.clear();

  var months = ['SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE', 'JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN'];
  var sessionsPerMonth = 9; // 9 créneaux par mois + TOT + %
  var extraCols = 2; // TOT + %

  var nameCol = 1; // Colonne A = Noms
  var firstDataCol = 2; // Données commencent en B

  // ---- Ligne 1 : Titre ----
  var totalCols = 1 + months.length * (sessionsPerMonth + extraCols);
  sh.getRange(1, 1, 1, totalCols).merge()
    .setValue('📅 SUIVI DE PRÉSENCE AUX ENTRAÎNEMENTS – Saison 2025/2026')
    .setBackground('#1a237e').setFontColor('#ffffff').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  // ---- Ligne 2 : Légende ----
  sh.getRange('A2').setValue('Légende :');
  sh.getRange('B2').setValue('P = Présent').setBackground('#c8e6c9');
  sh.getRange('D2').setValue('AJ = Absence Justifiée').setBackground('#fff9c4');
  sh.getRange('F2').setValue('AI = Absence Injustifiée').setBackground('#ffcdd2');
  sh.getRange('H2').setValue('B = Blessé/Malade').setBackground('#e1bee7');

  // ---- Ligne 3 : Mois (headers) ----
  // ---- Ligne 4 : En-têtes séances ----
  // ---- Ligne 5 : Noms joueurs ----

  var monthColors = [
    '#1565c0','#6a1b9a','#00695c','#bf360c',
    '#4527a0','#ad1457','#2e7d32','#f57f17',
    '#0277bd','#558b2f'
  ];

  var col = firstDataCol;
  for (var m = 0; m < months.length; m++) {
    var monthStart = col;
    var monthColor = monthColors[m % monthColors.length];

    // Ligne 3 : nom du mois centré sur (sessionsPerMonth + extraCols) colonnes
    sh.getRange(3, col, 1, sessionsPerMonth + extraCols).merge()
      .setValue(months[m])
      .setBackground(monthColor).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

    // Ligne 4 : en-têtes des séances
    for (var s = 1; s <= sessionsPerMonth; s++) {
      sh.getRange(4, col).setValue('Séance ' + s)
        .setBackground('#e3f2fd').setFontSize(8).setHorizontalAlignment('center').setFontWeight('bold');
      sh.setColumnWidth(col, 68);
      col++;
    }
    sh.getRange(4, col).setValue('TOT').setBackground('#bbdefb').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setColumnWidth(col, 45);
    col++;
    sh.getRange(4, col).setValue('%').setBackground('#bbdefb').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setColumnWidth(col, 45);
    col++;
  }

  // ---- Ligne 3, col A : "NOM / PRÉNOM" ----
  sh.getRange(3, 1, 2, 1).merge().setValue('NOM / PRÉNOM')
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setColumnWidth(1, 160);

  // ---- Lignes joueurs (5 à 34) ----
  var dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['P', 'AJ', 'AI', 'B'], true)
    .setAllowInvalid(false)
    .build();

  for (var row = 5; row <= 34; row++) {
    var bg = (row % 2 === 0) ? '#eceff1' : '#ffffff';

    // Référence au nom depuis Page 1
    sh.getRange(row, 1).setFormula("='Liste Joueurs'!B" + (row - 1) + "&\" \"&'Liste Joueurs'!C" + (row - 1))
      .setBackground('#cfd8dc').setFontWeight('bold');

    // Colonnes données
    var c2 = firstDataCol;
    for (var m2 = 0; m2 < months.length; m2++) {
      var sessionStart = c2;

      for (var s2 = 0; s2 < sessionsPerMonth; s2++) {
        sh.getRange(row, c2).setBackground(bg).setHorizontalAlignment('center');
        sh.getRange(row, c2).setDataValidation(dropdownRule);
        c2++;
      }

      // Colonne TOT : compte les "P"
      var startCell = columnLetter(sessionStart) + row;
      var endCell = columnLetter(sessionStart + sessionsPerMonth - 1) + row;
      sh.getRange(row, c2).setFormula('=COUNTIF(' + startCell + ':' + endCell + ',"P")')
        .setBackground('#bbdefb').setFontWeight('bold').setHorizontalAlignment('center');
      c2++;

      // Colonne % : présents / séances remplies
      var totCell = columnLetter(c2 - 1) + row;
      var countCell = columnLetter(sessionStart) + row + ':' + columnLetter(sessionStart + sessionsPerMonth - 1) + row;
      sh.getRange(row, c2).setFormula('=IFERROR(' + totCell + '/COUNTA(' + countCell + '),"")')
        .setBackground('#bbdefb').setNumberFormat('0%').setFontWeight('bold').setHorizontalAlignment('center');
      c2++;
    }
  }

  // ---- Ligne 35 : TOTAL séances ----
  sh.getRange(35, 1).setValue('TOTAL SÉANCES').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');

  sh.setFrozenRows(4);
  sh.setFrozenColumns(1);
}

// ─────────────────────────────────────────────
// PAGE 3 : CONVOCATIONS MATCHS
// ─────────────────────────────────────────────
function buildPage3(ss) {
  var sh = ss.getSheetByName('Convocations Matchs');
  sh.clear();

  var nbMatchs = 30;

  // Titre
  sh.getRange(1, 1, 1, nbMatchs + 1).merge()
    .setValue('📣 CONVOCATIONS MATCHS – Saison 2025/2026')
    .setBackground('#1a237e').setFontColor('#ffffff').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  // Légende
  sh.getRange('A2').setValue('Légende :');
  sh.getRange('B2').setValue('✅ Convoqué').setBackground('#c8e6c9');
  sh.getRange('D2').setValue('❌ Non convoqué').setBackground('#ffcdd2');
  sh.getRange('F2').setValue('⚠️ Convoqué absent').setBackground('#fff9c4');

  // En-têtes matchs
  sh.getRange(3, 1).setValue('NOM / PRÉNOM').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setColumnWidth(1, 160);

  for (var m = 1; m <= nbMatchs; m++) {
    var col = m + 1;
    sh.getRange(3, col).setValue('Match ' + m)
      .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setColumnWidth(col, 80);
  }

  // Dropdown convocation
  var convRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ Convoqué', '❌ Non convoqué', '⚠️ Convoqué absent'], true)
    .setAllowInvalid(false)
    .build();

  // Lignes joueurs
  for (var row = 4; row <= 33; row++) {
    var bg = (row % 2 === 0) ? '#eceff1' : '#ffffff';

    sh.getRange(row, 1).setFormula("='Liste Joueurs'!B" + (row - 2) + "&\" \"&'Liste Joueurs'!C" + (row - 2))
      .setBackground('#cfd8dc').setFontWeight('bold');

    for (var mc = 2; mc <= nbMatchs + 1; mc++) {
      sh.getRange(row, mc).setBackground(bg).setHorizontalAlignment('center').setDataValidation(convRule);
    }
  }

  // Ligne stats : taux convocation par match
  sh.getRange(34, 1).setValue('Nb convoqués').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(35, 1).setValue('% convoqués').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');

  for (var mc2 = 2; mc2 <= nbMatchs + 1; mc2++) {
    var colLetter = columnLetter(mc2);
    sh.getRange(34, mc2).setFormula('=COUNTIF(' + colLetter + '4:' + colLetter + '33,"✅ Convoqué")')
      .setBackground('#bbdefb').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(35, mc2).setFormula('=IFERROR(' + colLetter + '34/COUNTA(' + colLetter + '4:' + colLetter + '33),"")')
      .setBackground('#bbdefb').setNumberFormat('0%').setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}

// ─────────────────────────────────────────────
// PAGE 4 : FEUILLE DE MATCH
// ─────────────────────────────────────────────
function buildPage4(ss) {
  var sh = ss.getSheetByName('Feuille de Match');
  sh.clear();

  var nbMatchs = 20;
  var colsPerMatch = 4; // Tit/Rem | Minutes | Buts | Passes
  var headerRows = 6; // Lignes d'en-tête pour chaque match
  var nameCol = 1;
  var firstMatchCol = 2;

  // Titre
  var totalCols = 1 + nbMatchs * colsPerMatch;
  sh.getRange(1, 1, 1, totalCols).merge()
    .setValue('⚽ FEUILLE DE MATCH – Saison 2025/2026')
    .setBackground('#1b5e20').setFontColor('#ffffff').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  // Colonne nom
  sh.setColumnWidth(1, 160);

  // Pour chaque match : 6 lignes d'en-tête + colonnes
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Championnat', 'Amical', 'Coupe', 'Tournoi'], true)
    .setAllowInvalid(false).build();

  var domExtRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Domicile', 'Extérieur', 'Terrain neutre'], true)
    .setAllowInvalid(false).build();

  var roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Titulaire', 'Remplaçant', 'N/A'], true)
    .setAllowInvalid(false).build();

  for (var m = 0; m < nbMatchs; m++) {
    var startCol = firstMatchCol + m * colsPerMatch;
    var matchColors = ['#1b5e20','#1a237e','#4a148c','#b71c1c','#e65100','#006064','#33691e','#880e4f','#1565c0','#37474f'];
    var mc = matchColors[m % matchColors.length];

    // Ligne 2 : "Match N" sur 4 colonnes
    sh.getRange(2, startCol, 1, colsPerMatch).merge()
      .setValue('Match ' + (m + 1)).setBackground(mc).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

    // Ligne 3 : Date (merge 4 cols)
    sh.getRange(3, startCol, 1, colsPerMatch).merge()
      .setValue('').setBackground('#e8f5e9').setNumberFormat('dd/mm/yyyy');
    sh.getRange(3, startCol).setNote('Saisir la date du match');

    // Ligne 4 : Type (dropdown)
    sh.getRange(4, startCol, 1, colsPerMatch).merge()
      .setBackground('#f1f8e9').setDataValidation(typeRule).setHorizontalAlignment('center');
    sh.getRange(4, startCol).setNote('Type de match');

    // Ligne 5 : Dom/Ext (dropdown)
    sh.getRange(5, startCol, 1, colsPerMatch).merge()
      .setBackground('#f9fbe7').setDataValidation(domExtRule).setHorizontalAlignment('center');

    // Ligne 6 : Résultat
    sh.getRange(6, startCol, 1, colsPerMatch).merge()
      .setValue('').setBackground('#fff9c4').setHorizontalAlignment('center');
    sh.getRange(6, startCol).setNote('Ex: 3-1');

    // Sous-colonnes : en-têtes ligne 7
    var subHeaders = ['Tit/Rem', 'Min', 'Buts', 'PD'];
    var subColors = ['#a5d6a7','#c8e6c9','#dcedc8','#f0f4c3'];
    for (var sc = 0; sc < colsPerMatch; sc++) {
      sh.getRange(7, startCol + sc).setValue(subHeaders[sc])
        .setBackground(subColors[sc]).setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center');
      sh.setColumnWidth(startCol + sc, 55);
    }

    // Lignes joueurs (8 à 37)
    for (var row = 8; row <= 37; row++) {
      var bg = (row % 2 === 0) ? '#f9fbe7' : '#ffffff';
      sh.getRange(row, startCol).setBackground(bg).setDataValidation(roleRule).setHorizontalAlignment('center');
      sh.getRange(row, startCol + 1).setBackground(bg).setHorizontalAlignment('center'); // minutes
      sh.getRange(row, startCol + 2).setBackground(bg).setHorizontalAlignment('center'); // buts
      sh.getRange(row, startCol + 3).setBackground(bg).setHorizontalAlignment('center'); // passes
    }
  }

  // Colonne A - labels fixes
  // Lignes 2-7 : labels des en-têtes match
  sh.getRange(2, 1).setValue('MATCH').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(3, 1).setValue('Date').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(4, 1).setValue('Type').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(5, 1).setValue('Dom/Ext').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(6, 1).setValue('Score').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(7, 1).setValue('NOM / PRÉNOM').setBackground('#263238').setFontColor('#ffffff').setFontWeight('bold').setFontSize(8);

  // Noms joueurs
  for (var row2 = 8; row2 <= 37; row2++) {
    sh.getRange(row2, 1)
      .setFormula("='Liste Joueurs'!B" + (row2 - 6) + "&\" \"&'Liste Joueurs'!C" + (row2 - 6))
      .setBackground('#cfd8dc').setFontWeight('bold');
  }

  sh.setFrozenRows(7);
  sh.setFrozenColumns(1);
}

// ─────────────────────────────────────────────
// PAGE 5 : STATS GÉNÉRALES
// ─────────────────────────────────────────────
function buildPage5(ss) {
  var sh = ss.getSheetByName('Stats Générales');
  sh.clear();

  var headers = [
    'NOM / PRÉNOM',
    'Taux\nprésence\nentraîn.',
    'Nb matchs\nconvoqué',
    'Taux\nconvocation',
    'Matchs\ntitulaire',
    'Matchs\nremplaçant',
    'Minutes\njouées',
    'Buts\nmarqués',
    'Passes\ndécisives',
    'Indice IRE\n(Conv/Présence)'
  ];

  var colWidths = [160, 90, 90, 80, 80, 90, 80, 70, 80, 110];

  // Titre
  sh.getRange(1, 1, 1, headers.length).merge()
    .setValue('📊 STATISTIQUES GÉNÉRALES – Saison 2025/2026')
    .setBackground('#1a237e').setFontColor('#ffffff').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  // Ligne 2 : Légende IRE
  sh.getRange(2, 1, 1, headers.length).merge()
    .setValue('Indice IRE : 🟢 Entre 1 et 1,10 = Justement convoqué  |  🟠 < 1 = Sous-convoqué  |  🔴 > 1,10 = Sur-convoqué')
    .setBackground('#e8eaf6').setFontStyle('italic').setHorizontalAlignment('center');

  // En-têtes ligne 3
  for (var h = 0; h < headers.length; h++) {
    sh.getRange(3, h + 1).setValue(headers[h])
      .setBackground('#3949ab').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.setColumnWidth(h + 1, colWidths[h]);
  }
  sh.setRowHeight(3, 55);

  var nbMatchsP4 = 20;
  var colsPerMatch = 4;

  // Lignes joueurs 4-33
  for (var row = 4; row <= 33; row++) {
    var playerIdx = row - 2; // index pour pages 2, 3, 4
    var bg = (row % 2 === 0) ? '#e8eaf6' : '#ffffff';

    // Col A : Nom
    sh.getRange(row, 1).setFormula("='Liste Joueurs'!B" + (row - 2) + "&\" \"&'Liste Joueurs'!C" + (row - 2))
      .setBackground('#cfd8dc').setFontWeight('bold');

    // Col B : Taux présence entraînement (page 2)
    // Sur la page 2, les joueurs sont en ligne 5 à 34, la dernière colonne % se calcule
    // On va chercher toutes les colonnes % de chaque mois sur la page 2
    // Page 2 : col 1 = noms, puis par mois : 9 séances + TOT + %
    // Mois 0 : col 2 à 10 (séances), 11 (TOT), 12 (%)
    // Mois 1 : col 13 à 21, 22 (TOT), 23 (%)
    // etc.
    var monthPctCols = [];
    for (var mi = 0; mi < 10; mi++) {
      monthPctCols.push(2 + mi * 11 + 10); // col % du mois mi (0-indexed)
    }
    // Taux de présence global = moyenne des % mensuels non vides
    var p2Row = row - 1 + 4; // Ligne correspondante dans Page 2 (joueurs en 5-34)
    var p2Formulas = monthPctCols.map(function(c) {
      return "'Présence Entraînement'!" + columnLetter(c) + p2Row;
    });
    sh.getRange(row, 2)
      .setFormula('=IFERROR(AVERAGEIF({' + p2Formulas.join(',') + '},"<>"),"")')
      .setBackground(bg).setNumberFormat('0%').setHorizontalAlignment('center');

    // Col C : Nb matchs convoqué
    // Page 3 : joueurs en lignes 4 à 33 (joueur 1 = ligne 4), matchs en colonnes 2 à 31
    var p3Row = row - 2 + 3; // ligne dans page 3
    sh.getRange(row, 3)
      .setFormula("=COUNTIF('Convocations Matchs'!B" + p3Row + ":AE" + p3Row + ",\"✅ Convoqué\")")
      .setBackground(bg).setHorizontalAlignment('center');

    // Col D : Taux convocation = nb convoqué / nb matchs joués
    sh.getRange(row, 4)
      .setFormula('=IFERROR(C' + row + '/COUNTA(\'Convocations Matchs\'!B' + p3Row + ':AE' + p3Row + '),"")')
      .setBackground(bg).setNumberFormat('0%').setHorizontalAlignment('center');

    // Col E : Matchs titulaire (page 4)
    // Page 4 : joueurs en lignes 8 à 37, match m => col firstMatchCol + m*4
    // Titulaire/Remplaçant dans col startCol de chaque match
    var titFormulaParts = [];
    for (var mi2 = 0; mi2 < nbMatchsP4; mi2++) {
      var matchCol = 2 + mi2 * colsPerMatch;
      var p4Row = row - 4 + 7; // ligne dans page 4
      titFormulaParts.push("'Feuille de Match'!" + columnLetter(matchCol) + p4Row);
    }
    sh.getRange(row, 5)
      .setFormula('=COUNTIF({' + titFormulaParts.join(',') + '},"Titulaire")')
      .setBackground(bg).setHorizontalAlignment('center');

    // Col F : Matchs remplaçant
    sh.getRange(row, 6)
      .setFormula('=COUNTIF({' + titFormulaParts.join(',') + '},"Remplaçant")')
      .setBackground(bg).setHorizontalAlignment('center');

    // Col G : Minutes jouées
    var minFormulaParts = [];
    for (var mi3 = 0; mi3 < nbMatchsP4; mi3++) {
      var matchCol2 = 2 + mi3 * colsPerMatch + 1; // colonne minutes
      var p4Row2 = row - 4 + 7;
      minFormulaParts.push("'Feuille de Match'!" + columnLetter(matchCol2) + p4Row2);
    }
    sh.getRange(row, 7)
      .setFormula('=SUM(' + minFormulaParts.join(',') + ')')
      .setBackground(bg).setHorizontalAlignment('center');

    // Col H : Buts
    var butsFormulaParts = [];
    for (var mi4 = 0; mi4 < nbMatchsP4; mi4++) {
      var matchCol3 = 2 + mi4 * colsPerMatch + 2;
      var p4Row3 = row - 4 + 7;
      butsFormulaParts.push("'Feuille de Match'!" + columnLetter(matchCol3) + p4Row3);
    }
    sh.getRange(row, 8)
      .setFormula('=SUM(' + butsFormulaParts.join(',') + ')')
      .setBackground(bg).setHorizontalAlignment('center');

    // Col I : Passes décisives
    var pdFormulaParts = [];
    for (var mi5 = 0; mi5 < nbMatchsP4; mi5++) {
      var matchCol4 = 2 + mi5 * colsPerMatch + 3;
      var p4Row4 = row - 4 + 7;
      pdFormulaParts.push("'Feuille de Match'!" + columnLetter(matchCol4) + p4Row4);
    }
    sh.getRange(row, 9)
      .setFormula('=SUM(' + pdFormulaParts.join(',') + ')')
      .setBackground(bg).setHorizontalAlignment('center');

    // Col J : Indice IRE = Taux convocation / Taux présence
    sh.getRange(row, 10)
      .setFormula('=IFERROR(D' + row + '/B' + row + ',"")')
      .setBackground(bg).setNumberFormat('0.00').setHorizontalAlignment('center').setFontWeight('bold');
  }

  // Mise en forme conditionnelle IRE (colonne J)
  var ireRange = sh.getRange('J4:J33');
  var rules = [];

  // Rouge : > 1,10
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(1.10)
    .setBackground('#ffcdd2').setFontColor('#b71c1c').setBold(true)
    .setRanges([ireRange]).build());

  // Vert : entre 1 et 1,10
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 1.10)
    .setBackground('#c8e6c9').setFontColor('#1b5e20').setBold(true)
    .setRanges([ireRange]).build());

  // Orange : < 1
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(1)
    .setBackground('#ffe0b2').setFontColor('#e65100').setBold(true)
    .setRanges([ireRange]).build());

  sh.setConditionalFormatRules(rules);

  // Ligne de moyennes
  sh.getRange(34, 1).setValue('MOYENNE ÉQUIPE').setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  for (var col = 2; col <= 10; col++) {
    var colL = columnLetter(col);
    sh.getRange(34, col)
      .setFormula('=IFERROR(AVERAGE(' + colL + '4:' + colL + '33),"")')
      .setBackground('#263238').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center')
      .setNumberFormat(col === 2 || col === 4 ? '0%' : '0.00');
  }

  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}

// ─────────────────────────────────────────────
// UTILITAIRE : numéro de colonne → lettre(s)
// ─────────────────────────────────────────────
function columnLetter(col) {
  var letter = '';
  while (col > 0) {
    var mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
