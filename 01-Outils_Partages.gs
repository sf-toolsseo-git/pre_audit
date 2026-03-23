function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('🚀 Stratégie')
        .addItem('⚙️ Configuration', 'afficherFenetreConfiguration')
        .addItem('🧩 Clustering', 'afficherFenetrePreparationClustering')
        .addItem('📈 Pré-audit', 'ouvrirFenetrePreAudit')
        .addSeparator()
        .addItem('📊 Générer bilan', 'peuplerFeuilleBilan')
        .addSeparator()
        .addItem('💾 Sauvegarder pré-audit', 'sauvegarderPreAudit')
        .addItem('🔄 Restaurer', 'restaurerProprietesDepuisConfig')
        .addToUi();
}

function obtenirFeuilleParNom(nomFeuille) {
    Logger.log("Tentative d'accès à la feuille : " + nomFeuille);
    var classeur = SpreadsheetApp.getActiveSpreadsheet();
    var feuille = classeur.getSheetByName(nomFeuille);

    if (!feuille) {
        Logger.log("Feuille introuvable : " + nomFeuille);
        return null;
    }

    return feuille;
}

function reorganiserOnglets() {
    Logger.log("Début de la réorganisation des onglets.");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Ordre des feuilles souhaité : Bilan, GSC, Quick wins, Nouveaux mots-clés
    var ordreCible = ["Bilan - Devis", "Bilan - Full", "GSC", "Quick wins - Agile", "Quick wins - Full", "Nouveaux mots-clés - Contenu", "Nouveaux mots-clés - Full"];
    var position = 1;

    ordreCible.forEach(function(nomFeuille) {
        var feuille = ss.getSheetByName(nomFeuille);
        if (feuille) {
            try {
                // Déplacer la feuille à la position désirée (1-based index)
                ss.setActiveSheet(feuille);
                ss.moveActiveSheet(position);
                Logger.log("Onglet '" + nomFeuille + "' déplacé à la position " + position);
                position++;
            } catch (e) {
                // Ignore l'erreur si la feuille est déjà en position 1 (Bilan) et qu'on essaie de la bouger après la position 1
                Logger.log("Erreur lors du déplacement de l'onglet " + nomFeuille + ": " + e.toString());
            }
        }
    });
    Logger.log("Réorganisation des onglets terminée.");
}

function analyserCSV(contenu) {
    // détection simple du délimiteur (virgule ou point-virgule) basée sur la première ligne
    // nécessaire car Utilities.parseCsv demande un délimiteur explicite s'il n'est pas standard
    var premiereLigneEnd = contenu.indexOf("\n");
    if (premiereLigneEnd === -1) premiereLigneEnd = contenu.length;
    
    var premiereLigne = contenu.substring(0, premiereLigneEnd);
    var delimiteur = premiereLigne.indexOf(";") > -1 ? ";" : ",";

    // utilisation du parser natif de Google Apps Script
    // cette méthode est plus fiable : elle ne décale pas les colonnes si une cellule est vide
    // et gère nativement les guillemets (ex: "Informational, Transactional")
    try {
        var lignes = Utilities.parseCsv(contenu, delimiteur);
        return lignes;
    } catch (e) {
        Logger.log("Erreur critique parsing CSV : " + e.toString());
        return [];
    }
}

function peuplerFeuilleBilan() {
    Logger.log("Début de la génération des feuilles Bilan (Full & Devis).");
    
    // génération du bilan complet (potentiel max)
    // sources : Quick wins - Full & Nouveaux mots-clés - Full
    genererOngletBilanSpecifique(
        "Bilan - Full",
        "Quick wins - Full",
        "Nouveaux mots-clés - Full"
    );

    // génération du bilan devis (sélection calibrée)
    // sources : Quick wins - Agile & Nouveaux mots-clés - Contenu
    genererOngletBilanSpecifique(
        "Bilan - Devis",
        "Quick wins - Agile",
        "Nouveaux mots-clés - Contenu"
    );
    
    Logger.log("Génération des bilans terminée.");
}

function genererOngletBilanSpecifique(nomOngletCible, nomSourceQW, nomSourceNMC) {
    Logger.log("Mise à jour de l'onglet : " + nomOngletCible);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // récupération de la feuille existante
    var feuilleBilan = ss.getSheetByName(nomOngletCible);
    if (!feuilleBilan) {
        Logger.log("Erreur : La feuille '" + nomOngletCible + "' n'existe pas. Veuillez la créer avec les en-têtes.");
        return;
    }

    // 1. lecture des données sources
    var previsionsValues = [];
    var historiqueValues = [];
    var qwGainsValues = [];
    var nmcGainsValues = [];

    try {
        var feuilleGSC = ss.getSheetByName("GSC");
        var feuilleQW = ss.getSheetByName(nomSourceQW);
        var feuilleNMC = ss.getSheetByName(nomSourceNMC);
        
        // vérification de la feuille gsc (base obligatoire)
        if (!feuilleGSC) {
            Logger.log("Erreur : feuille 'GSC' introuvable.");
            return;
        }

        // lecture GSC (D2:F13 pour prévisions, B5:B16 pour historique N-1)
        previsionsValues = feuilleGSC.getRange("D2:F13").getValues();
        historiqueValues = feuilleGSC.getRange("B5:B16").getValues();

        // lecture gains QW (si la feuille existe)
        if (feuilleQW) {
            qwGainsValues = feuilleQW.getRange("F3:AB3").getValues();
        }

        // lecture gains NMC (si la feuille existe)
        if (feuilleNMC) {
            nmcGainsValues = feuilleNMC.getRange("F3:AB3").getValues();
        }

    } catch (e) {
        Logger.log("Erreur de lecture pour " + nomOngletCible + " : " + e.toString());
        return;
    }

    // vérification intégrité GSC
    if (previsionsValues.length !== 12 || historiqueValues.length !== 12) {
        Logger.log("Données GSC incomplètes ou incorrectes (12 mois attendus). Arrêt pour " + nomOngletCible);
        return;
    }

    var bilanData = [];
    var traficHistoriqueN1 = historiqueValues.map(function(row) { return row[0] || 0; });

    // nettoyage des données existantes (à partir de la ligne 2)
    var lastRow = feuilleBilan.getLastRow();
    if (lastRow >= 2) {
        feuilleBilan.getRange(2, 1, lastRow - 1, 12).clearContent();
    }

    // 2. construction des données (boucle sur 12 mois)
    for (var i = 0; i < 12; i++) {
        var ligne = [];
        
        // données de base
        var traficBase = previsionsValues[i][1] || 0; // colonne E de GSC -> colonne C du Bilan
        var traficHist = traficHistoriqueN1[i] || 0;

        // récupération des gains (index * 2 car cellules fusionnées dans sources)
        var gainQW = (qwGainsValues.length > 0 && qwGainsValues[0][i * 2]) ? qwGainsValues[0][i * 2] : 0;
        var gainNMC = (nmcGainsValues.length > 0 && nmcGainsValues[0][i * 2]) ? nmcGainsValues[0][i * 2] : 0;

        // calculs des totaux par scénario
        var totalQW = traficBase + gainQW;
        var totalNMC = traficBase + gainNMC;
        var totalMixte = traficBase + gainQW + gainNMC;

        // fonction utilitaire croissance
        var calcCroissance = function(prevu, histo) {
            if (histo === 0) return 0;
            return (prevu - histo) / histo;
        };

        // gestion de la date
        var dateValue = previsionsValues[i][0];
        var formattedDate = dateValue;
        if (Object.prototype.toString.call(dateValue) === '[object Date]') {
            formattedDate = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "MM/yyyy");
        }

        // remplissage de la ligne (colonnes A à L)
        ligne.push(formattedDate);                                  // A: date
        ligne.push(i + 1);                                          // B: mois relatif
        ligne.push(traficBase);                                     // C: trafic naturel estimé
        ligne.push(previsionsValues[i][2]);                         // D: croissance annuelle (base)
        
        ligne.push(gainQW);                                         // E: gain (optimisation)
        ligne.push(totalQW);                                        // F: total (optimisation)
        ligne.push(calcCroissance(totalQW, traficHist));            // G: croissance (optimisation)
        
        ligne.push(gainNMC);                                        // H: gain (contenu)
        ligne.push(totalNMC);                                       // I: total (contenu)
        ligne.push(calcCroissance(totalNMC, traficHist));           // J: croissance (contenu)
        
        ligne.push(totalMixte);                                     // K: total (mixte)
        ligne.push(calcCroissance(totalMixte, traficHist));         // L: croissance (mixte)

        bilanData.push(ligne);
    }

    // 3. écriture des données
    if (bilanData.length > 0) {
        feuilleBilan.getRange(2, 1, bilanData.length, 12).setValues(bilanData);
    }

    // 4. mise en forme des données (formats nombres uniquement)
    
    // formatage pourcentages
    feuilleBilan.getRange(2, 4, 12, 1).setNumberFormat("0%");         // D
    feuilleBilan.getRange(2, 7, 12, 1).setNumberFormat("+0%;-0%;0%"); // G
    feuilleBilan.getRange(2, 10, 12, 1).setNumberFormat("+0%;-0%;0%"); // J
    feuilleBilan.getRange(2, 12, 12, 1).setNumberFormat("+0%;-0%;0%"); // L

    // formatage nombres (milliers)
    feuilleBilan.getRange(2, 3, 12, 1).setNumberFormat("# ##0");      // C
    feuilleBilan.getRange(2, 5, 12, 2).setNumberFormat("# ##0");      // E, F
    feuilleBilan.getRange(2, 8, 12, 2).setNumberFormat("# ##0");      // H, I
    feuilleBilan.getRange(2, 11, 12, 1).setNumberFormat("# ##0");     // K
    
    // formatage texte (date)
    feuilleBilan.getRange(2, 1, 12, 1).setNumberFormat("@");

    // alignement général
    feuilleBilan.getRange(2, 1, 12, 12).setHorizontalAlignment("center");
    
    Logger.log("Onglet '" + nomOngletCible + "' mis à jour avec succès.");
}

function raccourcirNom(nom, maxLen) {
    if (!nom) return "";
    var texte = String(nom).trim();
    if (texte.length > maxLen) {
        return texte.substring(0, maxLen - 3) + "...";
    }
    return texte;
}

function extraireDomaineNettoye(str) {
    if (!str) return "";
    var d = str.toLowerCase().replace(/^(?:https?:\/\/)?(?:www\.)?/i, "").split('/')[0];
    return d;
}

function syncPropertiesToConfigSheet() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheetName = "CONFIG";
        var sheet = ss.getSheetByName(sheetName);
        
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
        }
        
        sheet.clear();
        sheet.hideSheet();
        var props = PropertiesService.getScriptProperties().getProperties();
        
        var groups = {
            "🛠️ GÉNÉRAL": [
                "PROJECT_TYPE", "CLIENT_NAME", "CLIENT_URL", 
                "CLIENT_STRENGTH", "CLIENT_BRAND"
            ],
            "📊 CONCURRENTS": [
                "COMP_NAME_1", 
                "COMPETITOR_1", "COMP_STRENGTH_1", "COMP_BRAND_1",
                "COMP_NAME_2", "COMPETITOR_2", "COMP_STRENGTH_2", "COMP_BRAND_2",
                "COMP_NAME_3", "COMPETITOR_3", "COMP_STRENGTH_3", "COMP_BRAND_3",
                "COMP_NAME_4", "COMPETITOR_4", "COMP_STRENGTH_4", "COMP_BRAND_4",
                "COMP_NAME_5", "COMPETITOR_5", "COMP_STRENGTH_5", "COMP_BRAND_5",
                "IS_MULTI_THEME"
            ],
            "🧠 IA & CLUSTERING": [
                "GEMINI_API_KEY", "URLS_CONTEXTE", "CONTEXTE_CLIENT"
            ],
            "📈 PRÉ-AUDIT": [
                "SLIDE_PRE_AUDIT_ID", "URL_REPONSES", "BRIEF_PRE_AUDIT", 
                "PREAUDIT_BESOIN_HTML", "PREAUDIT_BESOIN_TEXTE", "PREAUDIT_SOLUTION_HTML", "PREAUDIT_SOLUTION_TEXTE", 
                "ANALYSE_SEMRUSH_TITRE", "ANALYSE_SEMRUSH_KW_HTML", "ANALYSE_SEMRUSH_KW", "ANALYSE_SEMRUSH_TRAFIC_HTML", "ANALYSE_SEMRUSH_TRAFIC",
                "ANALYSE_THEME_TOP_TITRE", "ANALYSE_THEME_TOP", "ANALYSE_THEME_FLOP_TITRE", "ANALYSE_THEME_FLOP",
                "ANALYSE_SEGMENT_TOP_TITRE", "ANALYSE_SEGMENT_TOP", "ANALYSE_SEGMENT_FLOP_TITRE", "ANALYSE_SEGMENT_FLOP"
            ],
            "🎯 CTR": [
                "CTR_POS_1", "CTR_POS_2", "CTR_POS_3", "CTR_POS_4", "CTR_POS_5",
                "CTR_POS_6", "CTR_POS_7", "CTR_POS_8", "CTR_POS_9", "CTR_POS_10"
            ],
            "⚙️ AVANCÉ": [
                "COMPETITION_COEFF", "REF_POS", "REF_POS_S3", 
                "SEO_AGILE_JOURS", "SEO_AGILE_FREQ", "NB_CONTENUS_DEVISES", 
                "COEFF_S2", "COEFF_S3"
            ],
            "📦 AUTRES": []
        };
        var knownKeys = [];
        for (var g in groups) {
            knownKeys = knownKeys.concat(groups[g]);
        }
        
        for (var key in props) {
            if (key.indexOf('DATA_') === 0 || key.indexOf('_CACHE') === 0 || (props[key] && props[key].length > 40000)) {
                continue;
            }
            if (knownKeys.indexOf(key) === -1) {
                groups["📦 AUTRES"].push(key);
            }
        }
        
        var maxRows = 0;
        for (var g in groups) {
            if (groups[g].length + 1 > maxRows) {
                maxRows = groups[g].length + 1;
            }
        }
        
        var numGroups = Object.keys(groups).length;
        var grid = [];
        for (var r = 0; r < maxRows; r++) {
            grid[r] = new Array(numGroups * 3).fill("");
        }
        
        var gIdx = 0;
        for (var g in groups) {
            var cBase = gIdx * 3;
            grid[0][cBase] = g;
            grid[0][cBase + 1] = "Valeur";
            
            for (var i = 0; i < groups[g].length; i++) {
                var k = groups[g][i];
                grid[i + 1][cBase] = k;
                grid[i + 1][cBase + 1] = props[k] || "";
            }
            gIdx++;
        }
        
        if (maxRows > 0) {
            var range = sheet.getRange(1, 1, maxRows, numGroups * 3);
            range.setValues(grid);

            sheet.setFrozenRows(1);
            sheet.setHiddenGridlines(true);

            for (var i = 0; i < numGroups; i++) {
                var cBase = (i * 3) + 1;

                // Ligne d'en-tête : CLIP + taille police fixe pour bloquer l'auto-resize sur les emojis
                sheet.getRange(1, cBase, 1, 2)
                    .setBackground("#08133B").setFontColor("#FFFFFF").setFontWeight("bold")
                    .setHorizontalAlignment("center").setFontSize(10)
                    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

                // Colonne clés : CLIP obligatoire (certaines clés dépassent 180px et déclenchent le wrap)
                sheet.getRange(2, cBase, maxRows - 1, 1)
                    .setFontFamily("Courier New").setFontWeight("bold").setFontColor("#5f6368")
                    .setFontSize(10).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

                // Colonne valeurs
                sheet.getRange(2, cBase + 1, maxRows - 1, 1)
                    .setFontSize(10).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("top");

                sheet.setColumnWidth(cBase, 180);
                sheet.setColumnWidth(cBase + 1, 300);
                if (cBase + 2 <= numGroups * 3) {
                    sheet.setColumnWidth(cBase + 2, 30);
                }
            }

            // API Sheets v4 : seule méthode qui force pixelSize sans que le client remette "Ajuster aux données"
            SpreadsheetApp.flush();
            try {
                var token = ScriptApp.getOAuthToken();
                var requests = [{
                    updateDimensionProperties: {
                        range: {
                            sheetId: sheet.getSheetId(),
                            dimension: "ROWS",
                            startIndex: 0,
                            endIndex: maxRows
                        },
                        properties: { pixelSize: 21 },
                        fields: "pixelSize"
                    }
                }];
                var response = UrlFetchApp.fetch(
                    "https://sheets.googleapis.com/v4/spreadsheets/" + ss.getId() + ":batchUpdate",
                    {
                        method: "POST",
                        headers: { "Authorization": "Bearer " + token },
                        contentType: "application/json",
                        payload: JSON.stringify({ requests: requests }),
                        muteHttpExceptions: true
                    }
                );
                if (response.getResponseCode() !== 200) {
                    Logger.log("Sheets API v4 — erreur hauteurs : " + response.getContentText());
                }
            } catch (eV4) {
                Logger.log("Sheets API v4 — exception : " + eV4.message);
            }
        }
    } catch (e) {
        Logger.log("Erreur lors de la synchronisation vers l'onglet CONFIG : " + e.toString());
    }
}

function sauvegarderPreAudit() {
    Logger.log("Début de la sauvegarde du pré-audit.");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();

    var props = PropertiesService.getScriptProperties().getProperties();
    var clientName = props['CLIENT_NAME'];

    if (!clientName || clientName.trim() === "") {
        ui.alert("Erreur : Le nom du client n'est pas configuré. Veuillez le renseigner dans la configuration.");
        return;
    }

    var rootFolderId = "1ZNaNzKFGQ5_2NtsKKo1D6qwkaR-vDmbr";
    var targetFolder;
    try {
        targetFolder = DriveApp.getFolderById(rootFolderId);
    } catch (e) {
        Logger.log("Erreur accès dossier cible : " + e.toString());
        ui.alert("Erreur : Impossible d'accéder au dossier de sauvegarde.");
        return;
    }

    var fileName = clientName + " - Pré-audit";
    Logger.log("Copie du fichier sous le nom : " + fileName);

    try {
        var newSpreadsheet = ss.copy(fileName);
        var newFile = DriveApp.getFileById(newSpreadsheet.getId());
        newFile.moveTo(targetFolder);
        var newUrl = newSpreadsheet.getUrl();

        var configSheet = ss.getSheetByName("CONFIG");
        if (configSheet) {
            ss.deleteSheet(configSheet);
        }
        
        PropertiesService.getScriptProperties().deleteAllProperties();
        Logger.log("Propriétés du script et onglet CONFIG supprimés avec succès.");

        var htmlContent = '<div style="font-family: Arial, sans-serif; padding: 10px;">' +
            '<h3 style="margin-top: 0; color: #073763;">Sauvegarde pré-audit terminée</h3>' +
            '<p>Le fichier a été sauvegardé et réinitialisé.</p>' +
            '<p style="margin-top: 25px; text-align: center;"><a href="' + newUrl + '" target="_blank" style="background-color: #0b5394; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px;">Ouvrir la copie</a></p>' +
            '</div>';
        var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
            .setWidth(450)
            .setHeight(200);
            
        ui.showModalDialog(htmlOutput, 'Opération réussie');

    } catch (e) {
        Logger.log("Erreur critique lors de la sauvegarde pré-audit : " + e.toString());
        ui.alert("Une erreur est survenue : " + e.toString());
    }
}

function restaurerProprietesDepuisConfig() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("CONFIG");

    if (!configSheet) {
        ui.alert("Erreur : L'onglet CONFIG est introuvable. Restauration impossible.");
        return;
    }

    try {
        var data = configSheet.getDataRange().getValues();
        var propsToRestore = {};
        var count = 0;
        var keyRegex = /^[A-Z0-9_]{3,}$/;

        for (var r = 1; r < data.length; r++) {
            var row = data[r];
            for (var c = 0; c < row.length; c++) {
                var cellValue = String(row[c]).trim();
                if (keyRegex.test(cellValue) && (c + 1) < row.length) {
                    var val = row[c + 1] !== null ? String(row[c + 1]) : "";
                    propsToRestore[cellValue] = val;
                    count++;
                    c++; 
                }
            }
        }

        if (count > 0) {
            PropertiesService.getScriptProperties().setProperties(propsToRestore);
            ui.alert("Succès : " + count + " propriétés restaurées avec succès depuis l'onglet CONFIG.");
        } else {
            ui.alert("Information : Aucune propriété valide n'a été trouvée dans l'onglet CONFIG.");
        }

    } catch (e) {
        Logger.log("Erreur lors de la restauration depuis CONFIG : " + e.toString());
        ui.alert("Erreur lors de la restauration : " + e.toString());
    }
}