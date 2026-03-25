function ouvrirFenetrePreAudit() {
    var html = HtmlService.createHtmlOutputFromFile('06-preaudit')
        .setWidth(1400)
        .setHeight(1000)
        .setTitle('Pré-audit');
    SpreadsheetApp.getUi().showModelessDialog(html, '📈 Pré-audit');
}

function chargerConfigurationPreAudit() {
    Logger.log("=== DÉBUT : chargerConfigurationPreAudit ===");
    var props = PropertiesService.getScriptProperties().getProperties();
    var userProps = PropertiesService.getUserProperties().getProperties();
    var config = {
        clientName: props['CONF_CLIENT_NAME'] || "",
        clientUrl: props['CONF_CLIENT_URL'] || "",
        urlsContexte: props['PA_URLS_CONTEXTE'] || "",
        contexteClient: props['PA_CONTEXTE_CLIENT'] || "",
        slideId: props['PA_SLIDE_ID'] || "",
        brief: props['PA_BRIEF_CONSULTANT'] || "",
        urlReponses: props['PA_URL_FORM_REPONSES'] || "",
        contextePreaudit: props['PA_PROFILAGE_COMMERCIAL'] || "",
        besoinHtml: props['TAG_SLIDE_BESOIN_HTML'] || "",
        besoinTexte: props['TAG_SLIDE_BESOIN'] || "",
        solutionHtml: props['TAG_SLIDE_SOLUTION_HTML'] || "",
        solutionTexte: props['TAG_SLIDE_SOLUTION'] || "",
        titreSemrush: props['TITRE_SLIDE_SEMRUSH'] || "",
        analyseKwHtml: props['ANALYSE_SEMRUSH_MOT_CLE_HTML'] || "",
        analyseKwTexte: props['ANALYSE_SEMRUSH_MOT_CLE'] || "",
        analyseTraficHtml: props['ANALYSE_SEMRUSH_TRAFIC_HTML'] || "",
        analyseTraficTexte: props['ANALYSE_SEMRUSH_TRAFIC'] || "",
        activeTab: userProps['PREAUDIT_ACTIVE_TAB'] || "config",
        analyseThemeTopTitre: props['TITRE_SLIDE_THEMATIQUETOP_CLIENT'] || "",
        analyseThemeTop: props['ANALYSE_THEMATIQUETOP_CLIENT_1'] || "", // Using _1 as placeholder for full analysis
        analyseThemeFlopTitre: props['TITRE_SLIDE_THEMATIQUEFLOP_CLIENT'] || "",
        analyseThemeFlop: props['ANALYSE_THEMATIQUEFLOP_CLIENT_1'] || "",
        analyseSegmentTopTitre: props['TITRE_SLIDE_MCTOP_CLIENT'] || "",
        analyseSegmentTop: props['ANALYSE_MCTOP_CLIENT_1'] || "",
        analyseSegmentFlopTitre: props['TITRE_SLIDE_MCFLOP_CLIENT'] || "",
        analyseSegmentFlop: props['ANALYSE_MCFLOP_CLIENT_1'] || "",
        
        // Ajout des concurrents pour que le HTML puisse faire la liaison avec les favicons
        competitorName1: props['CONF_COMP_NAME_1'] || "",
        competitor1: props['CONF_COMP_URL_1'] || "",
        competitorName2: props['CONF_COMP_NAME_2'] || "",
        competitor2: props['CONF_COMP_URL_2'] || "",
        competitorName3: props['CONF_COMP_NAME_3'] || "",
        competitor3: props['CONF_COMP_URL_3'] || "",
        competitorName4: props['CONF_COMP_NAME_4'] || "",
        competitor4: props['CONF_COMP_URL_4'] || "",
        competitorName5: props['CONF_COMP_NAME_5'] || "",
        competitor5: props['CONF_COMP_URL_5'] || "",
        
        // Focus mot-clé
        focusKw: props['TARGET_KW'] || "",
        focusVol: props['TARGET_KW_SV'] || "",
        focusClientUrl: props['TARGET_URL_CLIENT'] || "",
        focusNoPage: props['TARGET_KW_CLIENT_POS'] === "-" ? "true" : "false",
        focusCompUrl: props['TARGET_URL_CONCURRENT'] || "",
        focusLocalisation: props['TARGET_LOCALISATION'] || "",
        
        // Nouveaux champs granulaires
        serpTitre1: props['SERP_ELEMENT_TITRE_1'] || "",
        serpDesc1: props['SERP_ELEMENT_DESC_1'] || "",
        serpSvg1: props['PLACEHOLDER_SERPELEMENT_1'] || "",
        serpTitre2: props['SERP_ELEMENT_TITRE_2'] || "",
        serpDesc2: props['SERP_ELEMENT_DESC_2'] || "",
        serpSvg2: props['PLACEHOLDER_SERPELEMENT_2'] || "",
        serpTitre3: props['SERP_ELEMENT_TITRE_3'] || "",
        serpDesc3: props['SERP_ELEMENT_DESC_3'] || "",
        serpSvg3: props['PLACEHOLDER_SERPELEMENT_3'] || "",
        serpTitre4: props['SERP_ELEMENT_TITRE_4'] || "",
        serpDesc4: props['SERP_ELEMENT_DESC_4'] || "",
        serpSvg4: props['PLACEHOLDER_SERPELEMENT_4'] || "",
        
        intentionTitre: props['FOCUS_INTENTION_TITRE'] || "",
        intentionDesc: props['FOCUS_INTENTION_DESC'] || "",
        intentionDescHtml: props['FOCUS_INTENTION_DESC'] || "", // On triche avec le texte brut pour le chargement s'il n'y a pas de HTML distinct stocké
        
        standard1: props['focus_standard_texte_1'] || "",
        standard1Html: props['focus_standard_texte_1'] || "",
        standard2: props['focus_standard_texte_2'] || "",
        standard2Html: props['focus_standard_texte_2'] || "",
        standard3: props['focus_standard_texte_3'] || "",
        standard3Html: props['focus_standard_texte_3'] || "",
        
        semantique1: props['focus_semantique_texte_1'] || "",
        semantique1Html: props['focus_semantique_texte_1'] || "",
        semantique2: props['focus_semantique_texte_2'] || "",
        semantique2Html: props['focus_semantique_texte_2'] || "",
        semantique3: props['focus_semantique_texte_3'] || "",
        semantique3Html: props['focus_semantique_texte_3'] || ""
    };
    Logger.log("=== FIN : chargerConfigurationPreAudit ===");
    return config;
}

function recupererDetailsMotCle(motCle) {
    Logger.log("=== DÉBUT : recupererDetailsMotCle ===");
    Logger.log("Mot-clé recherché : " + motCle);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Matrice");
    if (!sheet) {
        Logger.log("Erreur : l'onglet 'Matrice' est introuvable.");
        return { success: false, error: "Onglet Matrice introuvable." };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
        Logger.log("Erreur : la Matrice ne contient pas assez de données.");
        return { success: false, error: "Matrice vide." };
    }
    
    var headers = data[0];
    
    // Identifier les colonnes "URL "
    var urlClientIdx = -1;
    var urlCompIndices = [];
    var posCompIndices = [];
    
    var props = PropertiesService.getScriptProperties().getProperties();
    var clientName = props['CONF_CLIENT_NAME'] || "Client";
    
    for (var j = 0; j < headers.length; j++) {
        var h = String(headers[j]);
        if (h.indexOf("URL ") === 0) {
            var entName = h.substring(4).trim();
            if (entName === clientName || (urlClientIdx === -1 && entName === "Client")) {
                urlClientIdx = j;
            } else {
                urlCompIndices.push(j);
                // Trouver la position correspondante
                for (var c = 0; c < headers.length; c++) {
                    if (String(headers[c]) === "Pos " + entName) {
                        posCompIndices.push({ posIdx: c, urlIdx: j });
                        break;
                    }
                }
            }
        }
    }
    
    Logger.log("Index URL Client : " + urlClientIdx);
    Logger.log("Index Positions Concurrents : " + JSON.stringify(posCompIndices));
    
    var kwLower = motCle.toLowerCase();
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (String(row[0]).trim().toLowerCase() === kwLower) {
            Logger.log("Match trouvé à la ligne : " + (i + 1));
            var volume = row[1];
            var clientUrl = urlClientIdx > -1 ? String(row[urlClientIdx]).trim() : "";
            
            var bestCompUrl = "";
            var bestCompPos = 9999;
            
            for (var k = 0; k < posCompIndices.length; k++) {
                var pos = parseInt(row[posCompIndices[k].posIdx]);
                if (!isNaN(pos) && pos > 0 && pos < bestCompPos) {
                    bestCompPos = pos;
                    bestCompUrl = String(row[posCompIndices[k].urlIdx]).trim();
                }
            }
            
            if (bestCompUrl === "-" || bestCompUrl === "") bestCompUrl = "";
            
            var result = {
                success: true,
                volume: volume,
                clientUrl: clientUrl,
                compUrl: bestCompUrl
            };
            
            Logger.log("Résultat : " + JSON.stringify(result));
            return result;
        }
    }
    
    Logger.log("Aucun match trouvé pour le mot-clé.");
    return { success: false, error: "Mot-clé non trouvé." };
}

function sauvegarderConfigFocusMotCle(data) {
    Logger.log("=== DÉBUT : sauvegarderConfigFocusMotCle ===");
    try {
        var motCle = (data.kw || "").trim().toLowerCase();
        var clientUrl = (data.clientUrl || "").trim();
        var compUrl = (data.compUrl || "").trim();
        var localisation = (data.localisation || "").trim();
        
        // Valeur par défaut pour la localisation
        if (localisation === "") localisation = "france";

        var clientPos = "-";
        var compPos = "-";

        // Recherche des positions exactes dans la Matrice si le mot-clé existe
        if (motCle !== "") {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var sheet = ss.getSheetByName("Matrice");
            if (sheet) {
                var matriceData = sheet.getDataRange().getValues();
                var headers = matriceData[0];
                
                var targetRow = null;
                for (var i = 1; i < matriceData.length; i++) {
                    if (String(matriceData[i][0]).trim().toLowerCase() === motCle) {
                        targetRow = matriceData[i];
                        break;
                    }
                }
                
                if (targetRow) {
                    for (var j = 2; j < headers.length; j++) {
                        var h = String(headers[j]);
                        if (h.indexOf("URL ") === 0) {
                            var entityName = h.substring(4).trim();
                            var posIdx = -1;
                            
                            // Trouver la colonne "Pos" correspondante
                            for (var k = 2; k < headers.length; k++) {
                                if (String(headers[k]) === "Pos " + entityName) {
                                    posIdx = k;
                                    break;
                                }
                            }
                            
                            if (posIdx !== -1) {
                                var cellUrl = String(targetRow[j]).trim();
                                var cellPos = parseInt(targetRow[posIdx], 10);
                                
                                if (clientUrl !== "" && cellUrl === clientUrl && data.noPage !== "true") {
                                    if (!isNaN(cellPos) && cellPos > 0) clientPos = cellPos;
                                }
                                if (compUrl !== "" && cellUrl === compUrl) {
                                    if (!isNaN(cellPos) && cellPos > 0) compPos = cellPos;
                                }
                            }
                        }
                    }
                }
            }
        }

        // Écrasement si la case "Pas de page" est cochée
        if (data.noPage === "true") {
            clientPos = "-";
        }

        var props = PropertiesService.getScriptProperties();
        props.setProperties({
            'TARGET_KW': data.kw || "",
            'TARGET_KW_SV': data.vol || "",
            'TARGET_URL_CLIENT': data.clientUrl || "",
            'TARGET_KW_CLIENT_POS': String(clientPos),
            'TARGET_URL_CONCURRENT': data.compUrl || "",
            'TARGET_KW_CONCURRENT_POS': String(compPos),
            'TARGET_LOCALISATION': localisation
        });
        
        syncPropertiesToConfigSheet();
        
        Logger.log("Sauvegarde réussie. Pos Client: " + clientPos + " | Pos Concurrent: " + compPos);
        Logger.log("=== FIN : sauvegarderConfigFocusMotCle ===");
        return { success: true };
    } catch (e) {
        Logger.log("Erreur : " + e.message);
        throw new Error("Erreur lors de la sauvegarde du focus mot-clé : " + e.message);
    }
}

function recupererReponseFormulaire(urlForm) {
    if (!urlForm) return "";
    
    if (urlForm.indexOf("docs.google.com/forms") === -1) {
        throw new Error("L'URL fournie n'est pas un lien Google Forms valide.");
    }

    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var clientName = (props['CONF_CLIENT_NAME'] || "").toLowerCase().trim();
        var clientUrl = (props['CONF_CLIENT_URL'] || "").toLowerCase().replace(/^(?:https?:\/\/)?(?:www\.)?/i, "").split('/')[0].trim();
        
        Logger.log("=== RECHERCHE DE RÉPONSE FORMULAIRE (PAR ID) ===");
        Logger.log("Critères de filtrage : Nom [" + clientName + "] | Domaine [" + clientUrl + "]");

        // Extraction de l'ID du formulaire via Regex
        var formIdMatch = urlForm.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (!formIdMatch || !formIdMatch[1]) {
            throw new Error("Impossible d'extraire l'ID du formulaire depuis l'URL.");
        }
        var formId = formIdMatch[1];
        Logger.log("ID du formulaire extrait : " + formId);
        
        var form = FormApp.openById(formId);
        var allResponses = form.getResponses();
        
        if (allResponses.length === 0) {
            return "⚠️ Aucune réponse trouvée dans ce formulaire.";
        }

        var targetResponse = null;
        var questionCible = "Nom de votre entreprise & nom de domaine (ex : google.fr)";

        // Parcours des réponses de la plus récente à la plus ancienne
        for (var r = allResponses.length - 1; r >= 0; r--) {
            var itemResponses = allResponses[r].getItemResponses();
            for (var i = 0; i < itemResponses.length; i++) {
                var item = itemResponses[i];
                if (item.getItem().getTitle().trim() === questionCible) {
                    var val = String(item.getResponse()).toLowerCase().trim();
                    
                    var matchNom = clientName !== "" && (val.indexOf(clientName) > -1 || clientName.indexOf(val) > -1);
                    var matchUrl = clientUrl !== "" && (val.indexOf(clientUrl) > -1 || clientUrl.indexOf(val) > -1);

                    if (matchNom || matchUrl) {
                        targetResponse = allResponses[r];
                        Logger.log("Match trouvé à l'index " + r + " pour la valeur : " + val);
                        break;
                    }
                }
            }
            if (targetResponse) break;
        }

        if (!targetResponse) {
            Logger.log("Aucun match trouvé. Utilisation de la toute dernière réponse par défaut.");
            targetResponse = allResponses[allResponses.length - 1];
        }

        var finalItems = targetResponse.getItemResponses();
        var timestamp = targetResponse.getTimestamp();
        var dateFormatee = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        
        var resultat = "--- RÉPONSE DU " + dateFormatee + " ---\n\n";

        for (var j = 0; j < finalItems.length; j++) {
            var itemResponse = finalItems[j];
            var question = itemResponse.getItem().getTitle();
            var reponse = itemResponse.getResponse();

            if (Array.isArray(reponse)) {
                reponse = reponse.join(", ");
            }

            resultat += "Question : " + question + "\n";
            resultat += "Réponse : " + reponse + "\n";
            resultat += "-----------------------------------\n";
        }

        Logger.log("Extraction terminée avec succès.");
        return resultat;

    } catch (e) {
        Logger.log("Erreur dans recupererReponseFormulaire : " + e.message);
        throw new Error("Erreur lors de la récupération du formulaire : " + e.message);
    }
}

function sauvegarderConfigurationPreAudit(form) {
    Logger.log("=== DÉBUT : sauvegarderConfigurationPreAudit ===");
    var props = PropertiesService.getScriptProperties();
    props.setProperties({
        'CONF_CLIENT_NAME': form.clientName || "",
        'CONF_CLIENT_URL': form.clientUrl || "",
        'PA_URLS_CONTEXTE': form.urlsContexte || "",
        'PA_CONTEXTE_CLIENT': form.contexteClient || "",
        'PA_SLIDE_ID': form.slideId || "",
        'PA_BRIEF_CONSULTANT': form.brief || "",
        'PA_URL_FORM_REPONSES': form.urlReponses || "",
        'PA_PROFILAGE_COMMERCIAL': form.contextePreaudit || ""
    });
    syncPropertiesToConfigSheet();
    
    Logger.log("=== FIN : sauvegarderConfigurationPreAudit ===");
    return true;
}

function recupererArborescenceCluster() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Cluster");
    if (!sheet) throw new Error("L'onglet 'Cluster' n'existe pas. Veuillez le générer au préalable.");
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var arborescence = {};
    for (var i = 1; i < data.length; i++) {
        var theme = String(data[i][0]).trim();
        var subTheme = String(data[i][1]).trim();
        if (!theme || !subTheme) continue;
        if (!arborescence[theme]) arborescence[theme] = {};
        arborescence[theme][subTheme] = true;
    }
    var result = [];
    for (var t in arborescence) {
        result.push({ theme: t, subs: Object.keys(arborescence[t]).sort() });
    }
    result.sort(function(a, b) { return a.theme.localeCompare(b.theme); });
    return result;
}

function sauvegarderSelectionAnalyse(selection) {
    Logger.log("=== DÉBUT : sauvegarderSelectionAnalyse ===");
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperty('ANALYSE_SELECTION', JSON.stringify(selection));
        Logger.log("=== FIN : sauvegarderSelectionAnalyse ===");
        return true;
    } catch (e) {
        Logger.log("Erreur dans sauvegarderSelectionAnalyse : " + e.message);
        return false;
    }
}

function chargerSelectionAnalyse() {
    Logger.log("=== DÉBUT : chargerSelectionAnalyse ===");
    try {
        var props = PropertiesService.getScriptProperties();
        var data = props.getProperty('ANALYSE_SELECTION');
        Logger.log("=== FIN : chargerSelectionAnalyse ===");
        return data ? JSON.parse(data) : [];
    } catch (e) {
        Logger.log("Erreur dans chargerSelectionAnalyse : " + e.message);
        return [];
    }
}

function sauvegarderAnalysesEtatLieux(data) {
    Logger.log("=== DÉBUT : sauvegarderAnalysesEtatLieux ===");
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperties({
            'TITRE_SLIDE_THEMATIQUETOP_CLIENT': data.titreTopThematiques || "",
            'ANALYSE_THEMATIQUETOP_CLIENT_1': data.analyseTopThematiques || "",
            'TITRE_SLIDE_THEMATIQUEFLOP_CLIENT': data.titreFlopThematiques || "",
            'ANALYSE_THEMATIQUEFLOP_CLIENT_1': data.analyseFlopThematiques || "",
            'TITRE_SLIDE_MCTOP_CLIENT': data.titreTopSegments || "",
            'ANALYSE_MCTOP_CLIENT_1': data.analyseTopSegments || "",
            'TITRE_SLIDE_MCFLOP_CLIENT': data.titreFlopSegments || "",
            'ANALYSE_MCFLOP_CLIENT_1': data.analyseFlopSegments || ""
        });
        syncPropertiesToConfigSheet();
        Logger.log("=== FIN : sauvegarderAnalysesEtatLieux ===");
        return true;
    } catch (e) {
        Logger.log("Erreur lors de la sauvegarde des analyses IA : " + e.message);
        return false;
    }
}

function genererDiagnostic(selection) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCluster = ss.getSheetByName("Cluster");
    var sheetMatrice = ss.getSheetByName("Matrice");
    if (!sheetCluster) throw new Error("Onglet 'Cluster' introuvable.");
    if (!sheetMatrice) throw new Error("Onglet 'Matrice' introuvable.");
    if (!selection || selection.length === 0) throw new Error("Aucune thématique sélectionnée.");

    var props = PropertiesService.getScriptProperties().getProperties();

    // Récupération des CTR (stockés en pourcentage, ex : 28.5 pour 28,5 %)
    var ctrTable = [];
    for (var ci = 1; ci <= 10; ci++) {
        ctrTable.push((parseFloat(props['CTR_POS_' + ci]) || 0) / 100);
    }
    function computeTEC(vol, pos) {
        if (isNaN(pos) || pos <= 0 || pos > 10) return 0;
        return vol * ctrTable[pos - 1];
    }
    function computeTPM(vol) {
        return vol * (ctrTable[0] || 0);
    }

    var selectedSubs = {};
    selection.forEach(function(item) {
        selectedSubs[item.theme + "|" + item.sub] = true;
    });

    var clusterData = sheetCluster.getDataRange().getValues();
    var targetKeywords = new Map();
    for (var i = 1; i < clusterData.length; i++) {
        var theme = String(clusterData[i][0] || "").trim();
        var subTheme = String(clusterData[i][1] || "").trim();
        if (!selectedSubs[theme + "|" + subTheme]) continue;

        var intent = String(clusterData[i][5] || "").trim().toLowerCase();
        
        var mainKw = String(clusterData[i][3] || "").trim().toLowerCase();
        if (mainKw) {
            if (!targetKeywords.has(mainKw) || !targetKeywords.get(mainKw).isMain) {
                targetKeywords.set(mainKw, { theme: theme, sub: subTheme, intent: intent, isMain: true });
            }
        }
        
        var secKwStr = String(clusterData[i][10] || "").trim();
        if (secKwStr) {
            var secKws = secKwStr.split(/[\n,]+/);
            for (var j = 0; j < secKws.length; j++) {
                var sk = secKws[j].trim().toLowerCase();
                if (sk && !targetKeywords.has(sk)) {
                    targetKeywords.set(sk, { theme: theme, sub: subTheme, intent: intent, isMain: false });
                }
            }
        }
    }

    if (targetKeywords.size === 0) throw new Error("Aucun mot-clé trouvé pour cette sélection.");

    var matriceData = sheetMatrice.getDataRange().getValues();
    var headers = matriceData[0];
    var entities = [];
    for (var c = 2; c < headers.length; c++) {
        var h = String(headers[c]);
        if (h.indexOf("Pos ") === 0) {
            var name = h.substring(4).trim();
            var urlIdx = -1;
            for (var j = c + 1; j < headers.length; j++) {
                if (String(headers[j]) === "URL " + name) { urlIdx = j; break; }
            }
            entities.push({ name: name, posIdx: c, urlIdx: urlIdx });
        }
    }

    var clientName = props['CONF_CLIENT_NAME'] || "Client";
    if (clientName.trim() === "") clientName = "Client";
    var clientEntity = entities.find(function(e) { return e.name === clientName; });

    var kpis = {};
    entities.forEach(function(e) {
        kpis[e.name] = { posAll: 0, top3: 0, top10: 0, urls: new Set(), TEC: 0, TPM: 0 };
    });

    var themeStats = {};
    var intentStats = {
        transac: { kwCount: 0, top100: 0, top10: 0, TEC: 0, TPM: 0 },
        info:    { kwCount: 0, top100: 0, top10: 0, TEC: 0, TPM: 0 }
    };
    var acquis = [], gains = [], pertes = [], territoires = [];

    for (var i = 1; i < matriceData.length; i++) {
        var row = matriceData[i];
        var kw = String(row[0]).trim().toLowerCase();
        var vol = parseInt(row[1]) || 0;
        if (!targetKeywords.has(kw)) continue;

        var kwMeta = targetKeywords.get(kw);
        var tsKey = kwMeta.theme + " > " + kwMeta.sub;

        if (!themeStats[tsKey]) {
            themeStats[tsKey] = { kwCount: 0, volTotal: 0, top100: 0, top10: 0, top3: 0, TEC: 0, TPM: 0, DDT: 0, entityStats: {} };
            entities.forEach(function(e) { themeStats[tsKey].entityStats[e.name] = { TEC: 0 }; });
        }

        var kwTPM = computeTPM(vol);
        var clientPos = 999;
        var bestCompPos = 999;
        var bestCompName = "-";
        var compInTop10Count = 0;

        entities.forEach(function(e) {
            var pos = parseInt(row[e.posIdx]);
            var p = (!isNaN(pos) && pos > 0) ? pos : 999;
            var url = (e.urlIdx >= 0 && row[e.urlIdx]) ? String(row[e.urlIdx]).trim() : "";

            if (p <= 100) kpis[e.name].posAll++;
            if (p <= 3)   kpis[e.name].top3++;
            if (p <= 10)  kpis[e.name].top10++;
            if (url && url !== "-") kpis[e.name].urls.add(url);

            var eTEC = computeTEC(vol, p);
            kpis[e.name].TEC += eTEC;
            kpis[e.name].TPM += kwTPM;
            themeStats[tsKey].entityStats[e.name].TEC += eTEC;

            if (e.name === clientName) {
                clientPos = p;
            } else {
                if (p <= 10) compInTop10Count++;
                if (p < bestCompPos) { bestCompPos = p; bestCompName = e.name; }
            }
        });

        var clientTEC = computeTEC(vol, clientPos);
        var kwDDT = kwTPM - clientTEC;

        themeStats[tsKey].kwCount++;
        themeStats[tsKey].volTotal += vol;
        themeStats[tsKey].TPM += kwTPM;
        themeStats[tsKey].TEC += clientTEC;
        themeStats[tsKey].DDT += kwDDT;
        if (clientPos <= 100) themeStats[tsKey].top100++;
        if (clientPos <= 10)  themeStats[tsKey].top10++;
        if (clientPos <= 3)   themeStats[tsKey].top3++;

        var isTransac = kwMeta.intent.indexOf("transaction") > -1 || kwMeta.intent.indexOf("commercial") > -1 || kwMeta.intent === "t" || kwMeta.intent === "c";
        var isInfo    = kwMeta.intent.indexOf("information") > -1 || kwMeta.intent === "i";

        if (isTransac) {
            intentStats.transac.TPM += kwTPM;
            if (kwMeta.isMain) {
                intentStats.transac.kwCount++;
            }
        }
        if (isInfo) {
            intentStats.info.TPM += kwTPM;
            if (kwMeta.isMain) {
                intentStats.info.kwCount++;
            }
        }

        if (clientEntity) {
            if (isTransac) {
                intentStats.transac.TEC += clientTEC;
                if (kwMeta.isMain) {
                    if (clientPos <= 100) intentStats.transac.top100++;
                    if (clientPos <= 10)  intentStats.transac.top10++;
                }
            }
            if (isInfo) {
                intentStats.info.TEC += clientTEC;
                if (kwMeta.isMain) {
                    if (clientPos <= 100) intentStats.info.top100++;
                    if (clientPos <= 10)  intentStats.info.top10++;
                }
            }

            // Segmentation SWO (Uniquement les mots-clés principaux)
            if (kwMeta.isMain) {
                if (clientPos <= 10) {
                    acquis.push({ kw: kw, vol: vol, pos: clientPos, DDT: kwDDT });
                } else if (clientPos >= 11 && clientPos <= 20) {
                    gains.push({ kw: kw, vol: vol, pos: clientPos, DDT: kwDDT });
                } else if (clientPos > 20 && compInTop10Count >= 1) {
                    pertes.push({ kw: kw, vol: vol, pos: clientPos < 999 ? clientPos : null, DDT: kwDDT, bestCompName: bestCompName, bestCompPos: bestCompPos < 999 ? bestCompPos : null });
                }
                if (clientPos > 10 && bestCompPos > 10 && vol > 50) {
                    var bm = Math.min(clientPos, bestCompPos);
                    territoires.push({ kw: kw, vol: vol, DDT: kwDDT, bestPos: bm < 999 ? bm : null });
                }
            }
        }
    }

    // Calcul des KPI par entité
    var kpisArray = [];
    entities.forEach(function(e) {
        var eTPM = kpis[e.name].TPM;
        var eTEC = kpis[e.name].TEC;
        kpisArray.push({
            name: e.name,
            isClient: (e.name === clientName),
            posAll: kpis[e.name].posAll,
            top3: kpis[e.name].top3,
            top10: kpis[e.name].top10,
            urlsCount: kpis[e.name].urls.size,
            TEC: Math.round(eTEC),
            TPM: Math.round(eTPM),
            PdM: eTPM > 0 ? (eTEC / eTPM) * 100 : 0
        });
    });
    kpisArray.sort(function(a, b) {
        if (a.isClient) return -1;
        if (b.isClient) return 1;
        return b.top10 - a.top10;
    });

    // Construction du tableau des thématiques avec tous les ratios
    var themeArray = [];
    for (var k in themeStats) {
        var ts = themeStats[k];
        var PdM = ts.TPM > 0 ? (ts.TEC / ts.TPM) * 100 : 0;
        var TdP = ts.kwCount > 0 ? (ts.top100 / ts.kwCount) * 100 : 0;
        var entityPdM = {};
        entities.forEach(function(e) {
            var eTEC = ts.entityStats[e.name] ? ts.entityStats[e.name].TEC : 0;
            entityPdM[e.name] = ts.TPM > 0 ? (eTEC / ts.TPM) * 100 : 0;
        });
        themeArray.push({
            name: k, kwCount: ts.kwCount, volTotal: ts.volTotal,
            top100: ts.top100, top10: ts.top10, top3: ts.top3,
            TEC: Math.round(ts.TEC), TPM: Math.round(ts.TPM), DDT: Math.round(ts.DDT),
            PdM: PdM, TdP: TdP, entityPdM: entityPdM
        });
    }
    themeArray.sort(function(a, b) { return b.volTotal - a.volTotal; });

    // Identification top / flop thématique
    var topTheme = null, flopTheme = null;
    if (themeArray.length > 0) {
        var sortedPdM = themeArray.slice().sort(function(a, b) { return b.PdM - a.PdM || b.TdP - a.TdP; });
        topTheme = sortedPdM[0].name;
        var sortedDDT = themeArray.slice().sort(function(a, b) { return b.DDT - a.DDT || b.TPM - a.TPM || a.TdP - b.TdP; });
        flopTheme = sortedDDT[0].name;
    }

    // Ratios intentStats
    ['transac', 'info'].forEach(function(k) {
        var s = intentStats[k];
        s.PdM = s.TPM > 0 ? (s.TEC / s.TPM) * 100 : 0;
        s.TdP = s.kwCount > 0 ? (s.top100 / s.kwCount) * 100 : 0;
        s.TEC = Math.round(s.TEC);
        s.TPM = Math.round(s.TPM);
    });

    acquis.sort(function(a, b) { return b.vol - a.vol; });
    gains.sort(function(a, b) { return b.DDT - a.DDT; });
    pertes.sort(function(a, b) { return b.DDT - a.DDT; });
    territoires.sort(function(a, b) { return b.vol - a.vol; });

    return {
        kpis: kpisArray,
        themeStats: themeArray,
        intentStats: intentStats,
        topTheme: topTheme,
        flopTheme: flopTheme,
        acquis:      acquis.slice(0, 10),
        gains:       gains.slice(0, 10),
        pertes:      pertes.slice(0, 10),
        territoires: territoires.slice(0, 10)
    };
}

function analyserEvolutionSemrushIA(img1Base64, img1Mime, img2Base64, img2Mime, contexteClient) {
    Logger.log("=== DÉBUT : analyserEvolutionSemrushIA ===");
    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var apiKey = props['CONF_API_KEY_GEMINI'];
        
        if (!apiKey || apiKey.trim() === "") {
            throw new Error("Clé API Gemini introuvable dans la configuration générale.");
        }

        var promptText = "Tu es un expert SEO. Analyse ces deux captures d'écran issues de Semrush : la première montre l'évolution du nombre de mots-clés, la seconde montre l'évolution de l'estimation de trafic.\n\n" +
                         "Contraintes strictes :\n" +
                         "1. Concision maximale : 2 à 3 phrases maximum par bloc d'analyse.\n" +
                         "2. Tendances macro uniquement : concentre-toi sur les dynamiques globales (hausse, baisse, stagnation, pics).\n" +
                         "3. Zéro hallucination numérique : interdiction absolue de deviner, d'extrapoler ou de mentionner des chiffres à partir des axes visuels.\n" +
                         "4. Rédige une analyse directe et professionnelle à l'attention du prospect. Utilise le vouvoiement.\n" +
                         "5. Formatage important : encadre les termes ou expressions très importantes de ton analyse avec des astérisques simples (ex: le trafic est en *forte baisse* depuis l'année dernière). N'utilise pas le gras markdown standard (**).\n\n" +
                         "Règles typographiques obligatoires (français) à respecter à la lettre :\n" +
                         "- Majuscule uniquement au premier mot des puces et des phrases (sauf noms propres).\n" +
                         "- Pas de majuscule au premier mot à l'intérieur d'une parenthèse (sauf nom propre).\n" +
                         "- Pas de majuscule après les deux-points (:) car ce n'est pas une phrase complète.\n" +
                         "- Jours, mois et langues toujours en minuscule.\n" +
                         "- L'acronyme 'SEO' doit toujours être écrit en majuscules.\n\n" +
                         "Fournis ta réponse strictement au format JSON avec les clés exactes suivantes :\n" +
                         "- 'titre_slide' : un titre percutant résumant la tendance générale (sans astérisques).\n" +
                         "- 'analyse_mots_cles' : l'analyse visuelle de la courbe des mots-clés respectant les contraintes.\n" +
                         "- 'analyse_trafic' : l'analyse visuelle de la courbe de trafic respectant les contraintes.\n\n" +
                         "Profilage commercial du prospect (pour contextualiser ton analyse visuelle si pertinent) :\n" +
                         (contexteClient || "non renseigné.");

        var payload = {
            "contents": [
                {
                    "parts": [
                        {"text": promptText},
                        {
                            "inlineData": {
                                "mimeType": img1Mime,
                                "data": img1Base64
                            }
                        },
                        {
                            "inlineData": {
                                "mimeType": img2Mime,
                                "data": img2Base64
                            }
                        }
                    ]
                }
            ],
            "generationConfig": {
                "responseMimeType": "application/json"
            }
        };

        var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent";
        var options = {
            "method": "post",
            "contentType": "application/json",
            "headers": {
                "x-goog-api-key": apiKey
            },
            "payload": JSON.stringify(payload),
            "muteHttpExceptions": true
        };

        var response = UrlFetchApp.fetch(apiUrl, options);
        var json = JSON.parse(response.getContentText());

        if (response.getResponseCode() !== 200) {
            throw new Error(json.error ? json.error.message : "Erreur inattendue de l'API Gemini.");
        }

        if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
            var responseText = json.candidates[0].content.parts[0].text.trim();
            responseText = responseText.replace(/^```json\n/, '').replace(/\n```$/, '');
            Logger.log("=== FIN : analyserEvolutionSemrushIA (Succès) ===");
            return { success: true, jsonString: responseText };
        } else {
            throw new Error("L'API Gemini n'a renvoyé aucune analyse valide.");
        }

    } catch (e) {
        Logger.log("Erreur dans analyserEvolutionSemrushIA : " + e.message);
        return { success: false, error: e.message };
    }
}

function sauvegarderAnalyseEvolution(titre, texteKw, texteTrafic) {
    Logger.log("=== DÉBUT : sauvegarderAnalyseEvolution ===");
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperty('TITRE_SLIDE_SEMRUSH', titre || "");
        props.setProperty('ANALYSE_SEMRUSH_MOT_CLE', texteKw || "");
        props.setProperty('ANALYSE_SEMRUSH_TRAFIC', texteTrafic || "");
        Logger.log("=== FIN : sauvegarderAnalyseEvolution ===");
        return true;
    } catch (e) {
        throw new Error("Erreur lors de la sauvegarde : " + e.message);
    }
}

function sauvegarderOngletActif(tabName) {
    try {
        PropertiesService.getUserProperties().setProperty('PREAUDIT_ACTIVE_TAB', tabName);
        return true;
    } catch (e) {
        return false;
    }
}

function sauvegarderDonneesAnalyseGlobale(data) {
    Logger.log("=== DÉBUT : sauvegarderDonneesAnalyseGlobale ===");
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperties({
            'TAG_SLIDE_BESOIN_HTML':         data.besoinHtml || "",
            'TAG_SLIDE_BESOIN':              data.besoinTexte || "",
            'TAG_SLIDE_SOLUTION_HTML':       data.solutionHtml || "",
            'TAG_SLIDE_SOLUTION':            data.solutionTexte || "",
            'TITRE_SLIDE_SEMRUSH':           data.titreSemrush || "",
            'ANALYSE_SEMRUSH_MOT_CLE_HTML':  data.analyseKwHtml || "",
            'ANALYSE_SEMRUSH_MOT_CLE':       data.analyseKwTexte || "",
            'ANALYSE_SEMRUSH_TRAFIC_HTML':   data.analyseTraficHtml || "",
            'ANALYSE_SEMRUSH_TRAFIC':        data.analyseTraficTexte || "",
            'PLACEHOLDER_ANALYSE_SEMRUSH_MOT_CLE': "IMAGE",
            'PLACEHOLDER_ANALYSE_SEMRUSH_TRAFIC': "IMAGE"
        });
        
        Logger.log("Propriétés enregistrées, synchronisation en cours vers CONFIG...");
        syncPropertiesToConfigSheet();
        
        Logger.log("=== FIN : sauvegarderDonneesAnalyseGlobale ===");
        return true;
    } catch (e) {
        Logger.log("Erreur dans sauvegarderDonneesAnalyseGlobale : " + e.message);
        throw new Error("Erreur lors de la sauvegarde globale : " + e.message);
    }
}

function sauvegarderAnalysesFocus(data) {
    Logger.log("=== DÉBUT : sauvegarderAnalysesFocus ===");
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperties({
            'SERP_ELEMENT_TITRE_1': data.serpTitre1 || "",
            'SERP_ELEMENT_DESC_1': data.serpDesc1 || "",
            'PLACEHOLDER_SERPELEMENT_1': data.serpSvg1 || "",
            'SERP_ELEMENT_TITRE_2': data.serpTitre2 || "",
            'SERP_ELEMENT_DESC_2': data.serpDesc2 || "",
            'PLACEHOLDER_SERPELEMENT_2': data.serpSvg2 || "",
            'SERP_ELEMENT_TITRE_3': data.serpTitre3 || "",
            'SERP_ELEMENT_DESC_3': data.serpDesc3 || "",
            'PLACEHOLDER_SERPELEMENT_3': data.serpSvg3 || "",
            'SERP_ELEMENT_TITRE_4': data.serpTitre4 || "",
            'SERP_ELEMENT_DESC_4': data.serpDesc4 || "",
            'PLACEHOLDER_SERPELEMENT_4': data.serpSvg4 || "",
            
            'FOCUS_INTENTION_TITRE': data.intentionTitre || "",
            'FOCUS_INTENTION_DESC': data.intentionDesc || "",
            
            'focus_standard_texte_1': data.standard1 || "",
            'focus_standard_texte_2': data.standard2 || "",
            'focus_standard_texte_3': data.standard3 || "",
            
            'focus_semantique_texte_1': data.semantique1 || "",
            'focus_semantique_texte_2': data.semantique2 || "",
            'focus_semantique_texte_3': data.semantique3 || "",

            'FOCUS_GAP_TITRE_1': data.gapTitre1 || "",
            'FOCUS_GAP_DESC_1': data.gapDesc1 || "",
            'FOCUS_GAP_TITRE_2': data.gapTitre2 || "",
            'FOCUS_GAP_DESC_2': data.gapDesc2 || "",
            'FOCUS_GAP_TITRE_3': data.gapTitre3 || "",
            'FOCUS_GAP_DESC_3': data.gapDesc3 || "",

            'FOCUS_RECO_1': data.reco1 || "",
            'FOCUS_RECO_2': data.reco2 || "",
            'FOCUS_RECO_3': data.reco3 || "",
            'FOCUS_RECO_4': data.reco4 || ""
        });
        
        syncPropertiesToConfigSheet();
        Logger.log("Analyses Focus sauvegardées avec la nouvelle granularité.");
        Logger.log("=== FIN : sauvegarderAnalysesFocus ===");
        return true;
    } catch (e) {
        Logger.log("Erreur lors de la sauvegarde des analyses focus : " + e.message);
        return false;
    }
}

function lancerWorkflowSERP(data) {
    Logger.log("=== DÉBUT : lancerWorkflowSERP ===");
    Logger.log("Données reçues : " + JSON.stringify(data));
    
    // Dictionnaire de SVG (EXACTEMENT comme demandé)
    var dicoSVG = {
        "organique": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="2" y1="12" x2="22" y2="12"></line><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path></svg>',
        "paa": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="4" width="14" height="18" rx="2" ry="2"></rect><path d="M8 4V2h8v2"></path><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"></path><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>',
        "video": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polygon points="10 8 16 12 10 16 10 8"></polygon></svg>',
        "recherche": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>',
        "shopping": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="20" r="2"></circle><circle cx="19" cy="20" r="2"></circle><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"></path></svg>',
        "ads": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="M15.5 7.5a6 6 0 1 0 0 9"></path><line x1="6" y1="10" x2="14" y2="10"></line><line x1="6" y1="14" x2="14" y2="14"></line></svg>',
        "local": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>',
        "image": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><circle cx="8.5" cy="8.5" r="1.5"></circle><polyline points="21 15 16 10 5 21"></polyline></svg>',
        "featured": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>',
        "defaut": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="#1a73e8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>'
    };

    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var geminiApiKey = props['CONF_API_KEY_GEMINI'];
        var listeClesAPIStr = props['LISTE_CLES_API'];
        var contexteClient = props['PA_CONTEXTE_CLIENT'] || ""; // Préparation des données: Récupération du contexte client
        
        if (!geminiApiKey || geminiApiKey.trim() === "") {
            throw new Error("Clé API Gemini introuvable.");
        }
        
        var apiKeys = [];
        if (listeClesAPIStr) {
            try {
                var parsedKeys = JSON.parse(listeClesAPIStr);
                // Aplatit l'objet pour itérer plus facilement
                if (parsedKeys.serpapi) {
                    parsedKeys.serpapi.forEach(function(k) { apiKeys.push({type: 'serpapi', cle: k}); });
                }
                if (parsedKeys.serpstack) {
                    parsedKeys.serpstack.forEach(function(k) { apiKeys.push({type: 'serpstack', cle: k}); });
                }
            } catch (e) {
                Logger.log("Erreur de parsing LISTE_CLES_API : " + e.message);
            }
        }

        if (apiKeys.length === 0) throw new Error("Aucune clé API SERP configurée dans l'onglet Général.");
        var motCle = data.kw;
        var loc = data.localisation || "france";
        var urlClient = data.urlClient;
        var noPage = data.noPage;

        // ==========================================
        // ÉTAPE 1 : FETCH SERP (SERPAPI / SERPSTACK)
        // ==========================================
        Logger.log("Étape 1 : Récupération SERP pour '" + motCle + "'");
        var serpJson = fetchSerpData(motCle, loc, apiKeys);
        if (!serpJson) throw new Error("Impossible de récupérer la SERP (Quotas atteints ou erreur API).");
        var serpData = extractSerpData(serpJson);
        Logger.log("SERP Features détectées : " + serpData.features.join(", "));
        Logger.log("Top 10 URLs récupérées : " + serpData.urls.length);

        // ==========================================
        // ÉTAPE 2 : SCRAPING EN PARALLÈLE
        // ==========================================
        Logger.log("Étape 2 : Scraping des URLs du Top 10 + URL Client");
        var urlsToScrape = serpData.urls.slice(); // Copie du tableau
        
        // Si le client a une page cible qui n'est pas dans le top 10, on l'ajoute pour l'analyser
        if (!noPage && urlClient && urlsToScrape.indexOf(urlClient) === -1) {
            urlsToScrape.unshift(urlClient);
        }

        var scrapedPages = scrapeUrlsParallel(urlsToScrape);
        Logger.log("Scraping terminé. " + scrapedPages.length + " pages analysées.");

        // ==========================================
        // PRÉPARATION DU PAYLOAD POUR GEMINI
        // ==========================================
        var extractionData = {
            metadata: {
                keyword: motCle,
                location: loc,
                client_url: noPage ? "Page à créer" : urlClient
            },
            serp_features: serpData.features,
            pages: scrapedPages
        };

        Logger.log("Étape 3 : Création du prompt et appel API Gemini 2.5 Pro");
        
        var systemPrompt = `Tu es un expert stratège SEO. Ta mission est d'analyser des données brutes issues du scraping d'une SERP Google (top 10 concurrents + page cible du client) pour fournir un diagnostic clinique et un plan d'action stratégique sur-mesure.

            /// Protocole d'analyse (rigueur fiscale) ///

            1. Charte typographique (rigueur française absolue)
            - Titres, puces et labels : majuscule uniquement au premier mot (sauf noms propres).
            - Parenthèses : pas de majuscule au premier mot à l'intérieur.
            - Deux-points (:) : toujours un espace avant le deux-points, et aucune majuscule après (sauf nom propre).
            - Formatage visuel : utilise des astérisques simples *mot* pour mettre en gras les concepts clés de tes analyses pour faciliter la lecture en diagonale.

            2. Hiérarchie décisionnelle et objectifs (règle d'or : SERP > client)
            Ce que Google positionne dicte ce que tu dois recommander.
            - Mode création (si l'URL client est 'Page à créer') : ne critique pas la page client puisqu'elle n'existe pas. Projette la future page idéale basée sur les standards de la SERP.
            - Mode optimisation (si l'URL client existe) : audite la page cible existante du client par rapport aux exigences de la SERP.

            3. Intelligence commerciale et synergie (crucial)
            Ne traite pas le mot-clé cible en vase clos. Repère les expertises du client via le contexte client. Dans tes recommandations (et le gap business), propose des ponts intelligents (maillage interne, encarts de réassurance) entre le sujet analysé et l'offre globale du client.

            4. Lecture de l'environnement (SERP features)
            Examine les 'serp_features' détectées. Ce sont des signaux d'intention stricts :
            - Annonces (ads) : intention transactionnelle -> CTA visibles exigés.
            - PAA (questions) : intention informationnelle -> module FAQ exigé.
            - Pack local (maps) : proximité -> balisage local exigé.
            - Shopping : intention d'achat direct -> affichage prix/produits exigé.

            /// Format de sortie obligatoire ///
            Tu dois fournir ta réponse uniquement sous forme d'objet JSON valide, sans balise markdown autour.
            Structure stricte exigée :
            {
            "serp_elements": [
                {"titre": "Titre du bloc (ex : annonces Google Ads)", "description": "Ce que cela indique sur l'attente des utilisateurs.", "type_feature": "ads"}
                // exactement 4 objets. Clés autorisées pour type_feature : organique, paa, video, recherche, shopping, ads, local, image, featured, defaut.
            ],
            "intention": {
                "titre": "Typologie (ex : transactionnelle et commerciale)",
                "description": "Analyse concise de l'intention de l'utilisateur."
            },
            "standards": [
                "Standard structurel 1 (ex : présence de filtres à facettes ou tableau comparatif)",
                "Standard éditorial 2 (ex : ton académique ou vulgarisé)",
                "Standard UX 3 (ex : calculateur en ligne ou sommaire ancré)"
            ],
            "semantique": [
                "Axe lexical 1 à couvrir obligatoirement",
                "Axe lexical 2",
                "Axe lexical 3"
            ],
            "gap_analysis": [
                {"titre": "Format et UX", "description": "L'écart justifié entre le client et la SERP."},
                {"titre": "Profondeur éditoriale", "description": "Le manque à combler sur le contenu."},
                {"titre": "Business et synergie", "description": "L'opportunité manquée de vente ou maillage."}
                // exactement 3 objets. Si mode création, parle des "pièges à éviter".
            ],
            "recommandations": [
                "1. Action sur la structure de la page (Hn/UX).",
                "2. Action sur la profondeur du contenu.",
                "3. Action sur la conversion et réassurance.",
                "4. Action sur le maillage interne ou cross-sell (lien avec le contexte client)."
            ]
            }`;

                    var userPrompt = `[Contexte client] :
            ${contexteClient}

            [Données extraites de la SERP et du scraping] :
            ${JSON.stringify(extractionData)}`;

        var payload = {
            "system_instruction": {
                "parts": [{"text": systemPrompt}]
            },
            "contents": [
                {"parts": [{"text": userPrompt}]}
            ],
            "generationConfig": {
                "responseMimeType": "application/json"
            }
        };

        var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent";
        var options = {
            "method": "post",
            "contentType": "application/json",
            "headers": {
                "x-goog-api-key": geminiApiKey
            },
            "payload": JSON.stringify(payload),
            "muteHttpExceptions": true
        };

        Logger.log("Envoi de la requête à Gemini...");
        var response = UrlFetchApp.fetch(apiUrl, options);
        var jsonResponse = JSON.parse(response.getContentText());

        if (response.getResponseCode() !== 200) {
            throw new Error(jsonResponse.error ? jsonResponse.error.message : "Erreur inattendue de l'API Gemini.");
        }

        if (!jsonResponse.candidates || jsonResponse.candidates.length === 0 || !jsonResponse.candidates[0].content) {
            throw new Error("L'API Gemini n'a renvoyé aucune analyse valide.");
        }

        var responseText = jsonResponse.candidates[0].content.parts[0].text.trim();
        responseText = responseText.replace(/^```json\n/, '').replace(/\n```$/, '');

        Logger.log("Parsing de la réponse de Gemini...");
        var jsonGemini;
        try {
            jsonGemini = JSON.parse(responseText);
        } catch (e) {
            Logger.log("Erreur de parsing JSON Gemini : " + responseText);
            throw new Error("Le format JSON renvoyé par Gemini est invalide.");
        }

        // Mapping SVG : Boucle sur jsonGemini.serp_elements
        if (jsonGemini.serp_elements && Array.isArray(jsonGemini.serp_elements)) {
            for (var i = 0; i < jsonGemini.serp_elements.length; i++) {
                var el = jsonGemini.serp_elements[i];
                var featureKey = el.type_feature || "defaut";
                if (!dicoSVG[featureKey]) {
                    featureKey = "defaut";
                }
                el.svg_icon = dicoSVG[featureKey];
            }
        }

        Logger.log("=== FIN : lancerWorkflowSERP (Succès) ===");
        
        // Retour Front-End
        return {
            success: true,
            data: {
                elements_serp: jsonGemini.serp_elements || [],
                intention: jsonGemini.intention || {},
                standards: jsonGemini.standards || [],
                semantique: jsonGemini.semantique || [],
                gap_analysis: jsonGemini.gap_analysis || [],
                recommandations: jsonGemini.recommandations || []
            }
        };

    } catch(err) {
        Logger.log("Erreur dans lancerWorkflowSERP : " + err.message);
        return { success: false, error: err.message };
    }
}

function fetchSerpData(keyword, location, apiKeys) {
    // On priorise SerpApi (souvent plus complet sur les features)
    apiKeys.sort(function(a, b) {
        if (a.type === 'serpapi' && b.type !== 'serpapi') return -1;
        if (a.type !== 'serpapi' && b.type === 'serpapi') return 1;
        return 0;
    });

    for (var i = 0; i < apiKeys.length; i++) {
        var keyInfo = apiKeys[i];
        var url = "";
        
        if (keyInfo.type === 'serpapi') {
            url = "https://serpapi.com/search.json?q=" + encodeURIComponent(keyword) + "&hl=fr&gl=fr&google_domain=google.fr&api_key=" + keyInfo.cle;
            if (location) url += "&location=" + encodeURIComponent(location);
        } else if (keyInfo.type === 'serpstack') {
            url = "http://api.serpstack.com/search?access_key=" + keyInfo.cle + "&query=" + encodeURIComponent(keyword) + "&gl=fr&hl=fr&google_domain=google.fr";
            if (location) url += "&location=" + encodeURIComponent(location) + "&auto_location=0";
        }

        try {
            Logger.log("Tentative API avec : " + keyInfo.type);
            var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
            var json = JSON.parse(response.getContentText());

            if (json.error) {
                var errMsg = JSON.stringify(json.error).toLowerCase();
                if (errMsg.indexOf('quota') !== -1 || errMsg.indexOf('limit') !== -1) {
                    Logger.log("Quota dépassé pour " + keyInfo.type + ", passage à la suivante.");
                    continue;
                }
            }

            if (json.organic_results && json.organic_results.length > 0) {
                return json; // Succès
            }
        } catch (e) {
            Logger.log("Erreur SERP avec " + keyInfo.type + " : " + e.message);
        }
    }
    return null; // Toutes les clés ont échoué
}

function extractSerpData(json) {
    var urls = [];
    if (json.organic_results) {
        urls = json.organic_results.slice(0, 10).map(function(res) {
            return (res.link || res.url || "").replace(/\?srsltid.*/, '');
        }).filter(function(u) { return u !== ""; });
    }

    var features = [];
    if ((json.ads && json.ads.length > 0) || (json.advertisements && json.advertisements.length > 0)) features.push("Annonces (Google Ads)");
    if (json.answer_box || json.answer_box_list) features.push("Position 0 (Featured Snippet)");
    if (json.knowledge_graph) features.push("Knowledge Graph");
    if ((json.local_results && json.local_results.length > 0) || (json.places && json.places.length > 0)) features.push("Pack local (Maps)");
    if ((json.shopping_results && json.shopping_results.length > 0) || (json.inline_shopping && json.inline_shopping.length > 0) || (json.commercial_units && json.commercial_units.length > 0)) features.push("Bloc shopping");
    if (json.recipes_results && json.recipes_results.length > 0) features.push("Recettes");
    
    var paaCount = 0;
    if (json.related_questions) paaCount = json.related_questions.length;
    else if (json.people_also_ask) paaCount = json.people_also_ask.length;
    if (paaCount > 0) features.push("Autres questions posées (PAA)");
    
    if ((json.video_results && json.video_results.length > 0) || (json.inline_videos && json.inline_videos.length > 0)) features.push("Vidéos");
    if ((json.images_results && json.images_results.length > 0) || (json.inline_images && json.inline_images.length > 0)) features.push("Images");

    return { urls: urls, features: features };
}

function scrapeUrlsParallel(urls) {
    var userAgents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    ];

    // Création du tableau de requêtes pour UrlFetchApp.fetchAll
    var requests = urls.map(function(url) {
        return {
            url: url,
            method: "get",
            muteHttpExceptions: true,
            validateHttpsCertificates: false,
            followRedirects: true,
            headers: {
                'User-Agent': userAgents[Math.floor(Math.random() * userAgents.length)],
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7'
            }
        };
    });

    var results = [];
    try {
        // Exécution massive en parallèle (extrêmement rapide)
        var responses = UrlFetchApp.fetchAll(requests);
        
        responses.forEach(function(response, index) {
            var url = urls[index];
            if (response.getResponseCode() === 200) {
                var html = response.getContentText();
                var analysis = analyzePageContent(html);
                
                // Troncature pour ne pas exploser la fenêtre contextuelle (token limit) de Gemini
                var cleanText = analysis.cleaned_html;
                if (cleanText.length > 3000) cleanText = cleanText.substring(0, 3000) + "...";

                results.push({
                    url: url,
                    title: analysis.title,
                    structure_hn: analysis.structure_hn,
                    content_sample: cleanText
                });
            } else {
                results.push({ url: url, error: "HTTP " + response.getResponseCode() });
            }
        });
    } catch (e) {
        Logger.log("Erreur lors du fetchAll (Scraping) : " + e.message);
    }
    
    return results;
}

function analyzePageContent(html) {
    if (!html) return { title: "", structure_hn: [], cleaned_html: "" };
    
    // Sécurité : il faut s'assurer que la librairie Cheerio est bien activée dans le projet Apps Script
    if (typeof Cheerio === 'undefined') {
        return { title: "Erreur technique", structure_hn: [], cleaned_html: "La librairie Cheerio n'est pas chargée." };
    }

    var $ = Cheerio.load(html);
    var title = $('title').text().trim();
    
    // Nettoyage agressif des éléments parasites
    $('script, style, noscript, iframe, svg, link, meta, nav, footer, header, aside').remove();
    $('.cookie, .popup, .modal, .ad, .advertisement, .social-share, .login, .cart, .search-bar, .hidden, .d-none').remove();
    
    var structureHn = [];
    $('h1, h2, h3, h4').each(function() {
        var el = $(this);
        var txt = el.text().replace(/\s+/g, " ").trim();
        if (txt && txt.length > 2) {
            structureHn.push(el.prop("tagName").toLowerCase() + " : " + txt);
        }
    });

    var markdownParts = [];
    $('p, ul, ol, button, a').each(function() {
        var el = $(this);
        var tag = el.prop("tagName").toLowerCase();
        var text = el.text().replace(/\s+/g, " ").trim();

        if (text.length < 5) return;

        // Récupération des Boutons et Call-to-action
        if (tag === 'button' || (tag === 'a' && el.attr('class') && el.attr('class').toLowerCase().includes('btn'))) {
            if (text.length < 50) markdownParts.push("[CTA : " + text + "]");
        } 
        // Récupération du texte pertinent
        else if (tag === 'p' && text.length > 20) {
            markdownParts.push(text);
        } 
        // Récupération des listes
        else if (tag === 'ul' || tag === 'ol') {
            var items = [];
            el.find('li').each(function() {
                var liText = $(this).text().replace(/\s+/g, " ").trim();
                if (liText) items.push("- " + liText);
            });
            if (items.length > 0) markdownParts.push(items.join("\n"));
        }
    });

    // Fonction utilitaire pour décoder les entités HTML (si besoin)
    function decodeHtmlEntities(str) {
        return str.replace(/&#(\d+);/g, function(match, dec) { return String.fromCharCode(dec); })
                  .replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
    }

    return {
        title: decodeHtmlEntities(title),
        structure_hn: structureHn,
        cleaned_html: markdownParts.join("\n\n")
    };
}

function testRecuperationFormulaire() {
    var urlTest = "https://docs.google.com/forms/d/1ysuod7lKrpOjqb-wVPnfhauZmslPfnEXDdaRD4tAD2E/edit#response=ACYDBNjTwVbfs3pMEqdYAyukF4w8FbImmoocoXQSBk0VOnT0ZaL6tV8QLzUaYhHhptKRTWU";
    
    Logger.log("=== DÉBUT DU TEST DE RÉCUPÉRATION DU FORMULAIRE ===");
    Logger.log("URL testée : " + urlTest);
    
    try {
        var resultat = recupererReponseFormulaire(urlTest);
        Logger.log("=== RÉSULTAT OBTENU ===");
        Logger.log(resultat);
    } catch (e) {
        Logger.log("=== ERREUR DÉTECTÉE ===");
        Logger.log(e.message);
    }
    
    Logger.log("=== FIN DU TEST ===");
}

function genererProfilageCommercialIA(urlForm, brief, contexte) {
    Logger.log("=== DÉBUT : genererProfilageCommercialIA ===");
    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var apiKey = props['CONF_API_KEY_GEMINI'];
        
        if (!apiKey || apiKey.trim() === "") {
            throw new Error("Clé API Gemini introuvable. Veuillez configurer l'onglet Général.");
        }

        var reponsesForm = "Non disponible.";
        if (urlForm && urlForm.trim() !== "") {
            try {
                var extraction = recupererReponseFormulaire(urlForm);
                if (extraction && extraction.indexOf("⚠️ Aucune réponse") === -1) {
                    reponsesForm = extraction;
                }
            } catch (err) {
                Logger.log("Erreur lors de la récupération du formulaire : " + err.message);
                reponsesForm = "Erreur de lecture du formulaire : " + err.message;
            }
        }

        var promptStr = "Tu es un expert SEO et stratège en avant-vente. Ton objectif est d'analyser les sources d'informations disponibles pour dresser un profilage commercial redoutable du prospect. Ce contexte servira à orienter toutes les futures analyses pour l'aider à signer notre proposition d'accompagnement.\n\n" +
                        "RÈGLES DE PRIORITÉS (en cas de contradiction) :\n" +
                        "1. Source 1 : réponses du formulaire (priorité absolue, parole directe du prospect).\n" +
                        "2. Source 2 : brief commercial (notes du consultant).\n" +
                        "3. Source 3 : contexte site web (extraction automatisée).\n\n" +
                        "FORMAT DE SORTIE :\n" +
                        "Génère une synthèse stratégique en markdown, sans introduction ni conclusion. Utilise exclusivement des listes à puces sous les sections suivantes :\n" +
                        "- Profil et marché : (modèle économique, maturité perçue, positionnement)\n" +
                        "- Douleurs et frustrations : (problèmes actuels, échecs passés, ce qui le bloque)\n" +
                        "- Objectifs prioritaires : (kpi visés, attentes réelles au-delà du trafic, besoins vitaux)\n" +
                        "- Craintes et freins à l'achat : (objections potentielles, doutes sur le SEO ou le budget)\n" +
                        "- Angles d'attaque commerciaux : (arguments de réassurance à marteler, leviers psychologiques à activer)\n\n" +
                        "CONTRAINTES DE RÉDACTION :\n" +
                        "- Fusionne les informations complémentaires et supprime les redondances.\n" +
                        "- Rédige de manière incisive, orientée \"closing\" commercial.\n" +
                        "- Respect strict de la typographie : majuscule uniquement au premier mot des labels et débuts de ligne. Pas de majuscule après les deux-points.\n\n" +
                        "DONNÉES ENTRANTES :\n" +
                        "[SOURCE 1 - FORMULAIRE] :\n" + reponsesForm + "\n\n" +
                        "[SOURCE 2 - BRIEF] :\n" + (brief || "Non renseigné.") + "\n\n" +
                        "[SOURCE 3 - SITE WEB] :\n" + (contexte || "Non renseigné.");

        var payload = {
            "contents": [{
                "parts": [{"text": promptStr}]
            }]
        };

        var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent";
        var options = {
            "method": "post",
            "contentType": "application/json",
            "headers": {
                "x-goog-api-key": apiKey
            },
            "payload": JSON.stringify(payload),
            "muteHttpExceptions": true
        };

        var apiResponse = UrlFetchApp.fetch(apiUrl, options);
        var json = JSON.parse(apiResponse.getContentText());

        if (apiResponse.getResponseCode() !== 200) {
            throw new Error(json.error ? json.error.message : "Erreur inattendue de l'API Gemini.");
        }

        if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
            Logger.log("=== FIN : genererProfilageCommercialIA (Succès) ===");
            return { success: true, texte: json.candidates[0].content.parts[0].text.trim() };
        } else {
            throw new Error("L'API Gemini n'a renvoyé aucun texte.");
        }

    } catch (error) {
        Logger.log("Erreur dans genererProfilageCommercialIA : " + error.message);
        return { success: false, message: error.message };
    }
}

function genererSlideBesoinSolutionIA(contextePreaudit) {
    Logger.log("=== DÉBUT : genererSlideBesoinSolutionIA ===");
    var catalogueOffres = [
    "Audit et stratégie de positionnement : analyse concurrentielle, choix des mots-clés et plan d'action ciblé (mapping).",
    "SEO agile (accompagnement continu) : suivi mensuel/trimestriel, monitoring technique, reporting et recommandations d'optimisation.",
    "Accompagnement refonte : sécurisation SEO lors d'une création/refonte de site (cahier des charges, plan de redirection 301, recettes pré et post-lancement).",
    "Stratégie et rédaction de contenus SEO (clé en main) : planification éditoriale, rédaction experte et optimisation sémantique complète.",
    "Accompagnement éditorial (co-création) : fourniture des briefs/mots-clés et optimisation SEO de textes rédigés par le client."
    ];
    var texteOffres = "- " + catalogueOffres.join("\n- ");
    
    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var apiKey = props['CONF_API_KEY_GEMINI'];
        
        if (!apiKey || apiKey.trim() === "") {
            throw new Error("Clé API Gemini introuvable.");
        }

        var promptStr = "Tu es un expert SEO et stratège en avant-vente. À partir du profilage commercial fourni, tu dois extraire les arguments clés pour remplir une slide de présentation divisée en deux colonnes : 'Le constat (besoin)' et 'La réponse (solution)'.\n\n" +
                "Contraintes strictes de copywriting (crucial) :\n" +
                "1. Tu dois générer exactement 3 phrases pour le besoin, et exactement 3 phrases pour la solution. Pas une de plus, pas une de moins.\n" +
                "2. Style télégraphique : aucune phrase conversationnelle. N'utilise jamais les mots 'nous', 'notre', 'vous', 'votre', 'vos'.\n" +
                "3. Pour le 'Besoin', commence chaque puce obligatoirement par un nom commun (ex: Déficit, Nécessité, Manque, Volonté) ou un verbe à l'infinitif (ex: Définir, Structurer, Acquérir).\n" +
                "4. Pour la 'Solution', tu dois piocher uniquement dans ce catalogue d'offres : \n" + texteOffres + "\n" +
                "5. Pour la 'Solution', commence obligatoirement ta phrase par le NOM EXACT de l'offre choisie, suivi d'un espace, puis de deux-points, puis d'une phrase très courte de bénéfice.\n" +
                "6. Ne mets pas de tirets ou de puces textuelles dans la réponse JSON (le script s'en charge).\n\n" +
                "Règles typographiques obligatoires (français) à respecter à la lettre :\n" +
                "- Espacement : il faut TOUJOURS un espace avant les deux-points (ex: 'Audit SEO : pour...').\n" +
                "- Majuscule uniquement au premier mot des puces et des phrases (sauf noms propres).\n" +
                "- Pas de majuscule au premier mot à l'intérieur d'une parenthèse (sauf nom propre).\n" +
                "- Pas de majuscule après les deux-points (:) car ce n'est pas une phrase complète.\n" +
                "- Jours, mois et langues toujours en minuscule.\n" +
                "- L'acronyme 'SEO' doit toujours être écrit en majuscules.\n\n" +
                "Format de sortie attendu strictement en JSON :\n" +
                "{\n" +
                "  \"besoin\": [\"Phrase besoin 1\", \"Phrase besoin 2\", \"Phrase besoin 3\"],\n" +
                "  \"solution\": [\"Phrase solution 1\", \"Phrase solution 2\", \"Phrase solution 3\"]\n" +
                "}\n\n" +
                "Profilage commercial :\n" + contextePreaudit;

        var payload = {
            "contents": [{"parts": [{"text": promptStr}]}],
            "generationConfig": {
                "responseMimeType": "application/json"
            }
        };

        var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent";
        var options = {
            "method": "post",
            "contentType": "application/json",
            "headers": { "x-goog-api-key": apiKey },
            "payload": JSON.stringify(payload),
            "muteHttpExceptions": true
        };

        var apiResponse = UrlFetchApp.fetch(apiUrl, options);
        var json = JSON.parse(apiResponse.getContentText());

        if (apiResponse.getResponseCode() !== 200) {
            throw new Error(json.error ? json.error.message : "Erreur inattendue de l'API Gemini.");
        }

        if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
            var responseText = json.candidates[0].content.parts[0].text.trim();
            responseText = responseText.replace(/^```json\n/, '').replace(/\n```$/, '');
            Logger.log("=== FIN : genererSlideBesoinSolutionIA (Succès) ===");
            return { success: true, jsonString: responseText };
        } else {
            throw new Error("L'API Gemini n'a renvoyé aucune analyse valide.");
        }

    } catch (error) {
        Logger.log("Erreur dans genererSlideBesoinSolutionIA : " + error.message);
        return { success: false, message: error.message };
    }
}

function genererAnalyseTopFlopThemesIA(donneesTop, donneesFlop, contexteCommercial) {
    Logger.log("=== DÉBUT : genererAnalyseTopFlopThemesIA ===");
    Logger.log("Données Top reçues : " + donneesTop);
    Logger.log("Données Flop reçues : " + donneesFlop);
    Logger.log("Contexte commercial : " + contexteCommercial);

    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var apiKey = props['CONF_API_KEY_GEMINI'];
        
        if (!apiKey || apiKey.trim() === "") {
            Logger.log("Erreur : Clé API Gemini manquante.");
            throw new Error("Clé API Gemini introuvable.");
        }

        var promptStr = "Tu es un consultant SEO senior, pédagogue et fin stratège. Analyse simultanément ces deux tableaux JSON des thématiques du client : le 'Top 3' (trafic actuel) et le 'Flop 3' (manque à gagner). Ton objectif est d'en tirer un diagnostic clinique, factuel et orienté conseil, sans jamais tomber dans un discours commercial agressif.\n\n" +
                        "Lexique des métriques :\n" +
                        "- TEC (Trafic estimé client) : trafic actuel généré.\n" +
                        "- TPM (Trafic potentiel max) : trafic atteignable si 1ère position.\n" +
                        "- DDT (Déficit de trafic) : manque à gagner (TPM - TEC).\n\n" +
                        "Données JSON Top 3 :\n" + donneesTop + "\n\n" +
                        "Données JSON Flop 3 :\n" + donneesFlop + "\n\n" +
                        "Profilage commercial :\n" + (contexteCommercial || "Non renseigné.") + "\n\n" +
                        "RÈGLES DE RÉDACTION CRUCIALES (VERROUILLAGE SÉMANTIQUE) :\n" +
                        "1. NOMMAGE STRICT : tu dois obligatoirement utiliser le NOM EXACT des thématiques fournies (ex : « Sophrologie > Général »). Utilise exclusivement des guillemets français (« et ») pour les encadrer. N'utilise JAMAIS de guillemets simples ('), de doubles quotes standard (\") ou de backticks (`).\n" +
                        "2. CONTEXTE '> GÉNÉRAL' : si un segment se nomme '> Général', interprète-le comme des requêtes de notoriété ou de découverte (intentions informationnelles) et non comme l'intégralité de la thématique parente.\n" +
                        "3. STRUCTURE DES PUCES : commence chaque puce impérativement par le nom de la thématique concernée entre guillemets français pour ancrer l'analyse sur la donnée réelle.\n" +
                        "4. EXHAUSTIVITÉ : génère EXACTEMENT 3 puces pour le Top et EXACTEMENT 3 puces pour le Flop (une par thématique du JSON).\n" +
                        "5. PERSONNALISATION ET TON : Croise les données avec le profilage commercial. Adopte un ton d'expert SEO : sois rassurant sur le Top (protéger l'existant) et objectif sur le Flop (constater le décalage de visibilité sans dramatiser excessivement). Fuis les mots comme 'colossal', 'contre-attaquer', 'rentabilité immédiate'.\n\n" +
                        "CONTRAINTES DE STYLE ET TYPOGRAPHIE (HYPER-CONCISION OBLIGATOIRE) :\n" +
                        "- Concision extrême : MAXIMUM 2 lignes très courtes par puce. Rédige de façon chirurgicale. Élimine tout le blabla. Va droit au but : un constat, une recommandation, point final.\n" +
                        "- IMPÉRATIF VISUEL : encadre les concepts clés de ton analyse avec des astérisques simples pour le gras orange (ex : *effort de consolidation*, *potentiel inexploité*).\n" +
                        "- Règles FR : un espace obligatoire avant les deux-points (:). Pas de majuscule après les deux-points. Jours, mois et langues en minuscule. Acronymes (SEO, TEC, TPM, DDT) en majuscules.\n" +
                        "- TITRES (CRITIQUE) : Majuscule UNIQUEMENT au premier mot du titre. Tout le reste en minuscules (sauf noms propres). Exemple correct : 'Vos acquis actuels : une fondation solide à valoriser'. Exemple interdit : 'Vos Acquis Actuels : Une Fondation...'.\n\n" +
                        "Format de sortie STRICTEMENT JSON :\n" +
                        "{\n" +
                        "  \"titre_slide_top\": \"Titre valorisant vos acquis (règles de minuscules respectées)\",\n" +
                        "  \"analyse_top\": [\"La thématique « Nom exact » est un *socle solide* (491 visites). Notre objectif : combler le *déficit*.\", \"...\", \"...\"],\n" +
                        "  \"titre_slide_flop\": \"Titre d'alerte sur le manque à gagner (règles de minuscules respectées)\",\n" +
                        "  \"analyse_flop\": [\"Sur « Nom exact », vous accusez un *manque à gagner objectif* (DDT de 24 965). C'est un territoire à prioriser.\", \"...\", \"...\"]\n" +
                        "}\n" +
                        "Ne mets pas de tirets au début des phrases.";

        Logger.log("Envoi du payload à Gemini...");

        var payload = {
            "contents": [{"parts": [{"text": promptStr}]}],
            "generationConfig": { "responseMimeType": "application/json" }
        };

        var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent";
        var options = {
            "method": "post",
            "contentType": "application/json",
            "headers": { "x-goog-api-key": apiKey },
            "payload": JSON.stringify(payload),
            "muteHttpExceptions": true
        };

        var apiResponse = UrlFetchApp.fetch(apiUrl, options);
        var json = JSON.parse(apiResponse.getContentText());

        if (apiResponse.getResponseCode() !== 200) {
            Logger.log("Erreur API Gemini : " + apiResponse.getContentText());
            throw new Error(json.error ? json.error.message : "Erreur inattendue de l'API Gemini.");
        }

        if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
            var responseText = json.candidates[0].content.parts[0].text.trim();
            responseText = responseText.replace(/^```json\n/, '').replace(/\n```$/, '');
            Logger.log("Analyse générée avec succès.");
            Logger.log("=== FIN : genererAnalyseTopFlopThemesIA (Succès) ===");
            return { success: true, jsonString: responseText };
        } else {
            Logger.log("Erreur : l'API n'a pas renvoyé de contenu.");
            throw new Error("L'API Gemini n'a renvoyé aucune analyse valide.");
        }

    } catch (e) {
        Logger.log("ERREUR CRITIQUE : " + e.message);
        return { success: false, error: e.message };
    }
}

function genererAnalyseSegmentsIA(payloadTop, payloadFlop, contexteCommercial) {
    Logger.log("=== DÉBUT : genererAnalyseSegmentsIA ===");
    Logger.log("Payload Top reçu : " + payloadTop);
    Logger.log("Payload Flop reçu : " + payloadFlop);
    Logger.log("Contexte commercial : " + contexteCommercial);

    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var apiKey = props['CONF_API_KEY_GEMINI'];
        
        if (!apiKey || apiKey.trim() === "") {
            Logger.log("Erreur : Clé API Gemini manquante.");
            throw new Error("Clé API Gemini introuvable.");
        }

        var promptStr = "Tu es un consultant SEO senior, pédagogue et fin stratège. Ton objectif est de transformer des données SEO brutes (4 segments d'attaque) en une restitution d'audit percutante, factuelle et orientée conseil. Tu dois te baser sur le profilage psychologique et commercial du prospect fourni ci-dessous pour formuler des recommandations qui résonnent avec ses enjeux, sans jamais tomber dans un discours de vendeur agressif. Tu rédiges le contenu de 2 slides (une top et une flop) pour l'aider à prendre conscience de son potentiel inexploité.\n\n" +
                        "RÈGLE ABSOLUE DE PERSONNALISATION :\n" +
                        "Si un mot-clé présent dans les données JSON correspond à une thématique, un service, une douleur ou un objectif explicitement mentionné par le prospect dans le 'Profilage commercial', tu dois absolument le mettre en avant dans ton analyse pour montrer que cet audit répond directement à ses priorités métiers.\n\n" +
                        "Pour chaque slide, croise la donnée avec le profilage commercial selon cette méthode exacte :\n\n" +
                        "Slide 1 : 'top' (capitalisation et quick-wins - 3 puces au total)\n" +
                        "- Acquis stratégiques (top 3) : adopte une posture rassurante. Connecte ces mots-clés à ses 'objectifs' ou à sa 'maturité'. L'enjeu est factuel : protéger cet actif et cette notoriété historique face à l'inflation concurrentielle en ligne.\n" +
                        "- Gains immédiats (positions 4 à 20) : utilise ces mots-clés pour apaiser ses 'craintes et freins à l'achat' (ex : budget limité, niveau technique). Explique de façon experte qu'un effort SEO ciblé sur ces requêtes très proches du but permettra de consolider la visibilité rapidement sans nécessiter un budget colossal.\n\n" +
                        "Slide 2 : 'flop' (manque à gagner et vision - 3 puces au total)\n" +
                        "- Pertes de conquête (concurrents présents, pas lui) : appuie sur ses 'douleurs et frustrations' de façon objective. Cite les concurrents qui capitalisent sur son absence pour souligner le décalage entre sa légitimité métier et sa visibilité digitale. Adopte un ton de constat clinique (ex : 'visibilité captée par des pure-players' plutôt que 'ils vous volent des clients').\n" +
                        "- Territoires à prendre (océan bleu) : utilise ses 'angles d'attaque commerciaux'. Présente ces mots-clés (où la concurrence est plus faible) comme des espaces vierges parfaits pour évangéliser sa cible sur sa 'proposition de valeur unique' (ex : faire valoir ses propres certifications hors des sentiers battus).\n\n" +
                        "Adapte ton niveau de langage à sa 'maturité perçue' et ses 'craintes'. S'il est débutant, vulgarise et fuis le jargon. Fuis le vocabulaire commercial agressif. Tu es le spécialiste qui pose un diagnostic clair.\n\n" +
                        "CONTRAINTES DE STYLE ET TYPOGRAPHIE (HYPER-CONCISION OBLIGATOIRE) :\n" +
                        "- NOMMAGE STRICT : Utilise exclusivement des guillemets français (« et ») pour encadrer les mots-clés. N'utilise JAMAIS de guillemets simples ('), de doubles quotes standard (\") ou de backticks (`).\n" +
                        "- Concision extrême : MAXIMUM 2 lignes très courtes par puce. Rédige de façon chirurgicale. Élimine tout le blabla. Va droit au but : un constat, une recommandation.\n" +
                        "- Répartition stricte : génère exactement 3 puces pour la slide top, et 3 puces pour la slide flop.\n" +
                        "- Mise en valeur : encadre les concepts clés de ta recommandation avec des astérisques simples (ex : *effort SEO ciblé*, *manque à gagner objectif*).\n" +
                        "- Typographie FR : un espace obligatoire avant les deux-points (:). Pas de majuscule après les deux-points (sauf nom propre). Jours, mois et langues en minuscule. Acronymes (SEO, ROI) en majuscules.\n" +
                        "- TITRES : Majuscule UNIQUEMENT au premier mot du titre. Tout le reste en minuscules (sauf noms propres).\n" +
                        "- Vérité des données : cite les mots-clés exacts fournis dans les données JSON. N'invente aucun chiffre.\n\n" +
                        "Format de sortie JSON attendu :\n" +
                        "{\n" +
                        "  \"titre_slide_top\": \"Titre de la slide 1 (factuel et valorisant)\",\n" +
                        "  \"analyse_top\": [\"Puce 1 ultra courte sur le mot-clé « Mot-clé »\", \"Puce 2 ultra courte\", \"Puce 3 ultra courte\"],\n" +
                        "  \"titre_slide_flop\": \"Titre de la slide 2 (constat et opportunités)\",\n" +
                        "  \"analyse_flop\": [\"Puce 1 ultra courte sur le mot-clé « Mot-clé »\", \"Puce 2 ultra courte\", \"Puce 3 ultra courte\"]\n" +
                        "}\n\n" +
                        "Données JSON Top :\n" + payloadTop + "\n\n" +
                        "Données JSON Flop :\n" + payloadFlop + "\n\n" +
                        "Profilage commercial :\n" + (contexteCommercial || "Non renseigné.");

        Logger.log("Envoi du payload à Gemini pour les segments...");

        var payload = {
            "contents": [{"parts": [{"text": promptStr}]}],
            "generationConfig": { "responseMimeType": "application/json" }
        };

        var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent";
        var options = {
            "method": "post",
            "contentType": "application/json",
            "headers": { "x-goog-api-key": apiKey },
            "payload": JSON.stringify(payload),
            "muteHttpExceptions": true
        };

        var apiResponse = UrlFetchApp.fetch(apiUrl, options);
        var json = JSON.parse(apiResponse.getContentText());

        if (apiResponse.getResponseCode() !== 200) {
            Logger.log("Erreur API Gemini : " + apiResponse.getContentText());
            throw new Error(json.error ? json.error.message : "Erreur inattendue de l'API Gemini.");
        }

        if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
            var responseText = json.candidates[0].content.parts[0].text.trim();
            responseText = responseText.replace(/^```json\n/, '').replace(/\n```$/, '');
            
            Logger.log("Analyse segments générée avec succès.");
            Logger.log("=== RÉPONSE BRUTE DE L'IA ===");
            Logger.log(responseText);
            Logger.log("=== FIN : genererAnalyseSegmentsIA ===");
            return { success: true, jsonString: responseText };
        } else {
            Logger.log("Erreur : l'API n'a pas renvoyé de contenu.");
            throw new Error("L'API Gemini n'a renvoyé aucune analyse valide.");
        }

    } catch (e) {
        Logger.log("ERREUR CRITIQUE (Segments) : " + e.message);
        return { success: false, error: e.message };
    }
}
