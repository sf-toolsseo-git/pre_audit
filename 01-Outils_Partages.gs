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
    Logger.log("Synchronisation des propriétés vers l'onglet CONFIG...");
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
        
        var groups = [
            {
                name: "⚙️ CONFIG GLOBALE & CONTEXTE",
                keys: [
                    "CONF_PROJECT_TYPE", "CONF_API_KEY_GEMINI", "CONF_IS_MULTI_THEME", "PA_SLIDE_ID",
                    "CONF_CLIENT_NAME", "CONF_CLIENT_URL", "CONF_CLIENT_STRENGTH", "CONF_CLIENT_BRAND",
                    "CONF_COMP_NAME_1", "CONF_COMP_URL_1", "CONF_COMP_STRENGTH_1", "CONF_COMP_BRAND_1",
                    "CONF_COMP_NAME_2", "CONF_COMP_URL_2", "CONF_COMP_STRENGTH_2", "CONF_COMP_BRAND_2",
                    "CONF_COMP_NAME_3", "CONF_COMP_URL_3", "CONF_COMP_STRENGTH_3", "CONF_COMP_BRAND_3",
                    "CONF_COMP_NAME_4", "CONF_COMP_URL_4", "CONF_COMP_STRENGTH_4", "CONF_COMP_BRAND_4",
                    "CONF_COMP_NAME_5", "CONF_COMP_URL_5", "CONF_COMP_STRENGTH_5", "CONF_COMP_BRAND_5",
                    "CTR_POS_1", "CTR_POS_2", "CTR_POS_3", "CTR_POS_4", "CTR_POS_5",
                    "CTR_POS_6", "CTR_POS_7", "CTR_POS_8", "CTR_POS_9", "CTR_POS_10",
                    "PA_URL_FORM_REPONSES", "PA_URLS_CONTEXTE", "PA_BRIEF_CONSULTANT", "PA_CONTEXTE_CLIENT", "PA_PROFILAGE_COMMERCIAL",
                    "TARGET_KW", "TARGET_KW_SV", "TARGET_URL_CLIENT", "TARGET_KW_CLIENT_POS", "TARGET_URL_CONCURRENT", "TARGET_KW_CONCURRENT_POS", "TARGET_LOCALISATION"
                ]
            },
            {
                name: "🖼️ EXPORT SLIDES",
                keys: [
                    "TAG_SLIDE_BESOIN", "TAG_SLIDE_BESOIN_HTML", "TAG_SLIDE_SOLUTION", "TAG_SLIDE_SOLUTION_HTML",
                    "---BORDER---",
                    "TITRE_SLIDE_SEMRUSH", "ANALYSE_SEMRUSH_MOT_CLE", "ANALYSE_SEMRUSH_MOT_CLE_HTML", "ANALYSE_SEMRUSH_TRAFIC", "ANALYSE_SEMRUSH_TRAFIC_HTML", "PLACEHOLDER_ANALYSE_SEMRUSH_MOT_CLE", "PLACEHOLDER_ANALYSE_SEMRUSH_TRAFIC",
                    "---BORDER---",
                    "MOTCLE_CLIENT_GLOBAL", "MOTCLE_CLIENT_TOP10", "MOTCLE_CLIENT_TOP3", "MOTCLE_CLIENT_URL", "MOTCLE_CLIENT_TRANSAC", "MOTCLE_CLIENT_TRANSAC_TOP10", "MOTCLE_CLIENT_TRANSAC_PCT", "JAUGE_TRANSAC_TOP10", "MOTCLE_CLIENT_INFO", "MOTCLE_CLIENT_INFO_TOP10", "MOTCLE_CLIENT_INFO_PCT", "JAUGE_INFO_TOP10",
                    "---BORDER---",
                    "TITRE_SLIDE_THEMATIQUETOP_CLIENT", "ANALYSE_THEMATIQUETOP_CLIENT_1", "ANALYSE_THEMATIQUETOP_CLIENT_1_HTML", "ANALYSE_THEMATIQUETOP_CLIENT_2", "ANALYSE_THEMATIQUETOP_CLIENT_2_HTML", "ANALYSE_THEMATIQUETOP_CLIENT_3", "ANALYSE_THEMATIQUETOP_CLIENT_3_HTML", "top_thm_client_1", "top_thm_client_top10_1", "top_thm_client_tec_1", "top_thm_client_tpm_1", "top_thm_client_ddt_1", "top_thm_client_2", "top_thm_client_top10_2", "top_thm_client_tec_2", "top_thm_client_tpm_2", "top_thm_client_ddt_2", "top_thm_client_3", "top_thm_client_top10_3", "top_thm_client_tec_3", "top_thm_client_tpm_3", "top_thm_client_ddt_3",
                    "---BORDER---",
                    "TITRE_SLIDE_THEMATIQUEFLOP_CLIENT", "ANALYSE_THEMATIQUEFLOP_CLIENT_1", "ANALYSE_THEMATIQUEFLOP_CLIENT_1_HTML", "ANALYSE_THEMATIQUEFLOP_CLIENT_2", "ANALYSE_THEMATIQUEFLOP_CLIENT_2_HTML", "ANALYSE_THEMATIQUEFLOP_CLIENT_3", "ANALYSE_THEMATIQUEFLOP_CLIENT_3_HTML", "flop_thm_client_1", "flop_thm_client_flop10_1", "flop_thm_client_tec_1", "flop_thm_client_tpm_1", "flop_thm_client_ddt_1", "flop_thm_client_2", "flop_thm_client_flop10_2", "flop_thm_client_tec_2", "flop_thm_client_tpm_2", "flop_thm_client_ddt_2", "flop_thm_client_3", "flop_thm_client_flop10_3", "flop_thm_client_tec_3", "flop_thm_client_tpm_3", "flop_thm_client_ddt_3",
                    "---BORDER---",
                    "TITRE_SLIDE_MCTOP_CLIENT", "ANALYSE_MCTOP_CLIENT_1", "ANALYSE_MCTOP_CLIENT_1_HTML", "ANALYSE_MCTOP_CLIENT_2", "ANALYSE_MCTOP_CLIENT_2_HTML", "ANALYSE_MCTOP_CLIENT_3", "ANALYSE_MCTOP_CLIENT_3_HTML", "top_mc_client_1", "top_mc_client_vol_1", "top_mc_client_ddt_1", "top_mc_client_pos_1", "qw_mc_client_1", "qw_mc_client_vol_1", "qw_mc_client_ddt_1", "qw_mc_client_pos_1", "top_mc_client_2", "top_mc_client_vol_2", "top_mc_client_ddt_2", "top_mc_client_pos_2", "qw_mc_client_2", "qw_mc_client_vol_2", "qw_mc_client_ddt_2", "qw_mc_client_pos_2", "top_mc_client_3", "top_mc_client_vol_3", "top_mc_client_ddt_3", "top_mc_client_pos_3", "qw_mc_client_3", "qw_mc_client_vol_3", "qw_mc_client_ddt_3", "qw_mc_client_pos_3", "top_mc_client_4", "top_mc_client_vol_4", "top_mc_client_ddt_4", "top_mc_client_pos_4", "qw_mc_client_4", "qw_mc_client_vol_4", "qw_mc_client_ddt_4", "qw_mc_client_pos_4", "top_mc_client_5", "top_mc_client_vol_5", "top_mc_client_ddt_5", "top_mc_client_pos_5", "qw_mc_client_5", "qw_mc_client_vol_5", "qw_mc_client_ddt_5", "qw_mc_client_pos_5",
                    "---BORDER---",
                    "TITRE_SLIDE_MCFLOP_CLIENT", "ANALYSE_MCFLOP_CLIENT_1", "ANALYSE_MCFLOP_CLIENT_1_HTML", "ANALYSE_MCFLOP_CLIENT_2", "ANALYSE_MCFLOP_CLIENT_2_HTML", "ANALYSE_MCFLOP_CLIENT_3", "ANALYSE_MCFLOP_CLIENT_3_HTML", "pc_mc_client_1", "pc_mc_client_vol_1", "pc_mc_client_ddt_1", "pc_mc_conc_pos_1", "tap_mc_client_1", "tap_mc_client_vol_1", "tap_mc_client_ddt_1", "tap_mc_conc_pos_1", "pc_mc_client_2", "pc_mc_client_vol_2", "pc_mc_client_ddt_2", "pc_mc_conc_pos_2", "tap_mc_client_2", "tap_mc_client_vol_2", "tap_mc_client_ddt_2", "tap_mc_conc_pos_2", "pc_mc_client_3", "pc_mc_client_vol_3", "pc_mc_client_ddt_3", "pc_mc_conc_pos_3", "tap_mc_client_3", "tap_mc_client_vol_3", "tap_mc_client_ddt_3", "tap_mc_conc_pos_3", "pc_mc_client_4", "pc_mc_client_vol_4", "pc_mc_client_ddt_4", "pc_mc_conc_pos_4", "tap_mc_client_4", "tap_mc_client_vol_4", "tap_mc_client_ddt_4", "tap_mc_conc_pos_4", "pc_mc_client_5", "pc_mc_client_vol_5", "pc_mc_client_ddt_5", "pc_mc_conc_pos_5", "tap_mc_client_5", "tap_mc_client_vol_5", "tap_mc_client_ddt_5", "tap_mc_conc_pos_5",
                    "---BORDER---",
                    "TITRE_SLIDE_CONCURRENCE", "NOM_CLIENT", "VALEUR_TOP10_CLIENT", "VALEUR_PAGES_CLIENT", "PLACEHOLDER_LOGO_CLIENT", "NOM_LEADER", "VALEUR_TOP10_LEADER", "VALEUR_PAGES_LEADER", "PLACEHOLDER_LOGO_LEADER", "NOM_COMP1", "VALEUR_TOP10_COMP1", "VALEUR_PAGES_COMP1", "PLACEHOLDER_LOGO_COMP1", "NOM_COMP2", "VALEUR_TOP10_COMP2", "VALEUR_PAGES_COMP2", "PLACEHOLDER_LOGO_COMP2", "NOM_COMP3", "VALEUR_TOP10_COMP3", "VALEUR_PAGES_COMP3", "PLACEHOLDER_LOGO_COMP3", "NOM_COMP4", "VALEUR_TOP10_COMP4", "VALEUR_PAGES_COMP4", "PLACEHOLDER_LOGO_COMP4",
                    "---BORDER---",
                    "SERP_ELEMENT_TITRE_1", "SERP_ELEMENT_DESC_1", "PLACEHOLDER_SERPELEMENT_1", "SERP_ELEMENT_TITRE_2", "SERP_ELEMENT_DESC_2", "PLACEHOLDER_SERPELEMENT_2", "SERP_ELEMENT_TITRE_3", "SERP_ELEMENT_DESC_3", "PLACEHOLDER_SERPELEMENT_3", "SERP_ELEMENT_TITRE_4", "SERP_ELEMENT_DESC_4", "PLACEHOLDER_SERPELEMENT_4", "FOCUS_INTENTION_TITRE", "FOCUS_INTENTION_DESC", "focus_standard_texte_1", "focus_semantique_texte_1", "focus_standard_texte_2", "focus_semantique_texte_2", "focus_standard_texte_3", "focus_semantique_texte_3",
                    "---BORDER---",
                    "FOCUS_GAP_TITRE_1", "FOCUS_GAP_DESC_1", "FOCUS_GAP_TITRE_2", "FOCUS_GAP_DESC_2", "FOCUS_GAP_TITRE_3", "FOCUS_GAP_DESC_3", "FOCUS_RECO_1", "FOCUS_RECO_2", "FOCUS_RECO_3", "FOCUS_RECO_4"
                ]
            },
            {
                name: "🛠️ TECHNIQUE",
                keys: [
                    "TECH_URL_CIBLE", "TECH_SITEMAP", "TECH_TYPE_PAGE", "TECH_URL_PAGE_MERE", "TECH_URL_PAGINEES", "TECH_URL_FILTRE", "TECH_IS_MULTILINGUE", "TECH_LANGUE_CIBLE", "TECH_PAYS_CIBLE",
                    "---BORDER---",
                    "TECH_CRAWL_STATUS_CODE", "TECH_CRAWL_TTFB_MS", "TECH_CRAWL_TTFB_SCORE",
                    "TECH_CRAWL_ROBOTS_BLOCKED", "TECH_CRAWL_ROBOTS", "TECH_CRAWL_HREFLANGS",
                    "TECH_CRAWL_FIRST_LINK_URL", "TECH_CRAWL_FIRST_LINK_ANCHOR",
                    "TECH_CRAWL_LINKS_TOTAL", "TECH_CRAWL_LINKS_INTERNAL", "TECH_CRAWL_LINKS_EXTERNAL",
                    "TECH_CRAWL_LINKS_200", "TECH_CRAWL_LINKS_3XX", "TECH_CRAWL_LINKS_4XX", "TECH_CRAWL_LINKS_5XX",
                    "TECH_CRAWL_PAGI_MERE_BODY", "TECH_CRAWL_PAGI_MERE_HEAD", "TECH_CRAWL_PAGI_P2_BODY", "TECH_CRAWL_PAGI_P2_HEAD", "TECH_CRAWL_PAGI_ERREUR_LIEN",
                    "---BORDER---",
                    "TECH_INDEX_SITEMAP_PRESENT", "TECH_INDEX_URL_IN_SITEMAP", "TECH_INDEX_META_ROBOTS", "TECH_INDEX_CANONICAL", "TECH_INDEX_PAGI_META_ROBOTS", "TECH_INDEX_PAGI_CANONICAL",
                    "---BORDER---",
                    "TECH_POS_TITLE", "TECH_POS_TITLE_HAS_KW", "TECH_POS_H1", "TECH_POS_H1_HAS_KW", "TECH_POS_HN", "TECH_POS_SCHEMA",
                    "---BORDER---",
                    "CRAWL_CHECK_1", "CRAWL_CONTENT_1", "CRAWL_CHECK_2", "CRAWL_CONTENT_2", "CRAWL_CHECK_3", "CRAWL_CONTENT_3",
                    "INDEX_CHECK_1", "INDEX_CONTENT_1", "INDEX_CHECK_2", "INDEX_CONTENT_2", "INDEX_CHECK_3", "INDEX_CONTENT_3",
                    "POS_CHECK_1", "POS_CONTENT_1", "POS_CHECK_2", "POS_CONTENT_2", "POS_CHECK_3", "POS_CONTENT_3"
                ]
            },
            {
                name: "📦 AUTRES",
                keys: []
            }
        ];
        var knownKeys = [];
        for (var i = 0; i < groups.length; i++) {
            for (var j = 0; j < groups[i].keys.length; j++) {
                if (groups[i].keys[j] !== "---BORDER---") {
                    knownKeys.push(groups[i].keys[j]);
                }
            }
        }
        
        for (var key in props) {
            if (key.indexOf('DATA_') === 0 || key.indexOf('_CACHE') === 0 || (props[key] && props[key].length > 40000)) {
                continue;
            }
            // Correction : exclusion stricte des clés purement numériques
            if (!isNaN(key)) {
                continue;
            }
            if (knownKeys.indexOf(key) === -1) {
                groups[groups.length - 1].keys.push(key);
            }
        }
        
        var maxRows = 0;
        for (var i = 0; i < groups.length; i++) {
            if (groups[i].keys.length + 1 > maxRows) {
                maxRows = groups[i].keys.length + 1;
            }
        }
        
        var numGroups = groups.length;
        var grid = [];
        for (var r = 0; r < maxRows; r++) {
            grid[r] = new Array(numGroups * 3).fill("");
        }
        
        var borderRanges = [];
        for (var gIdx = 0; gIdx < numGroups; gIdx++) {
            var g = groups[gIdx];
            var cBase = gIdx * 3;
            grid[0][cBase] = g.name;
            grid[0][cBase + 1] = "Valeur";
            
            var rowOffset = 1;
            for (var i = 0; i < g.keys.length; i++) {
                var k = g.keys[i];
                if (k === "---BORDER---") {
                    borderRanges.push({row: rowOffset, colBase: cBase});
                } else {
                    grid[rowOffset][cBase] = k;
                    grid[rowOffset][cBase + 1] = props[k] || "";
                    rowOffset++;
                }
            }
        }
        
        if (maxRows > 0) {
            var range = sheet.getRange(1, 1, maxRows, numGroups * 3);
            range.setValues(grid);

            sheet.setFrozenRows(1);
            sheet.setHiddenGridlines(true);

            for (var i = 0; i < numGroups; i++) {
                var cBase = (i * 3) + 1;
                // En-têtes centrés
                sheet.getRange(1, cBase, 1, 2)
                    .setBackground("#08133B").setFontColor("#FFFFFF").setFontWeight("bold")
                    .setHorizontalAlignment("center").setFontSize(10)
                    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
                // Alignement à gauche pour les clés
                sheet.getRange(2, cBase, maxRows - 1, 1)
                    .setFontFamily("Courier New").setFontWeight("bold").setFontColor("#5f6368")
                    .setFontSize(10).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
                    .setHorizontalAlignment("left");
                // Alignement à gauche pour les valeurs
                sheet.getRange(2, cBase + 1, maxRows - 1, 1)
                    .setFontSize(10).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("top")
                    .setHorizontalAlignment("left");
                sheet.setColumnWidth(cBase, 350);
                sheet.setColumnWidth(cBase + 1, 350);
                if (cBase + 2 <= numGroups * 3) {
                    sheet.setColumnWidth(cBase + 2, 30);
                }
            }

            // Appliquer les bordures pour le groupe Slides
            for (var b = 0; b < borderRanges.length; b++) {
                var br = borderRanges[b];
                var cell1 = sheet.getRange(br.row, br.colBase + 1);
                var cell2 = sheet.getRange(br.row, br.colBase + 2);
                cell1.setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
                cell2.setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
            }

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
            } catch (eV4) {
                Logger.log("Sheets API v4 — exception : " + eV4.message);
            }
        }
    } catch (e) {
        Logger.log("Erreur lors de la synchronisation vers l'onglet CONFIG : " + e.toString());
    }
    Logger.log("=== FIN : syncPropertiesToConfigSheet ===");
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