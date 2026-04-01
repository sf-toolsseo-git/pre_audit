function onOpen() {
    Logger.log("Début de la fonction onOpen");
    var ui = SpreadsheetApp.getUi();
    Logger.log("Création du menu Stratégie en cours");
    ui.createMenu('🚀 Stratégie')
        .addItem('⚙️ Configuration', 'afficherFenetreConfiguration')
        .addItem('🧩 Clustering', 'afficherFenetrePreparationClustering')
        .addItem('📈 Pré-audit', 'ouvrirFenetrePreAudit')
        .addSeparator()
        .addItem('📊 Générer bilan', 'peuplerFeuilleBilan')
        .addSeparator()
        .addItem('💾 Sauvegarder pré-audit', 'sauvegarderPreAudit')
        .addItem('🔄 Restaurer', 'restaurerProprietesDepuisConfig')
        .addItem('🔑 Clés API', 'afficherFenetreClesAPI')
        .addToUi();
    Logger.log("Fin de la fonction onOpen");
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
        // Correction : la clé doit commencer par une lettre ou un underscore pour éviter de prendre les chiffres purs
        var keyRegex = /^[A-Z_][A-Z0-9_]{2,}$/;
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