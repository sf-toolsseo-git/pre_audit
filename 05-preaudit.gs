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
        analyseThemeTop: props['ANALYSE_THEMATIQUETOP_CLIENT_1'] || "",
        analyseThemeFlopTitre: props['TITRE_SLIDE_THEMATIQUEFLOP_CLIENT'] || "",
        analyseThemeFlop: props['ANALYSE_THEMATIQUEFLOP_CLIENT_1'] || "",
        analyseSegmentTopTitre: props['TITRE_SLIDE_MCTOP_CLIENT'] || "",
        analyseSegmentTop: props['ANALYSE_MCTOP_CLIENT_1'] || "",
        analyseSegmentFlopTitre: props['TITRE_SLIDE_MCFLOP_CLIENT'] || "",
        analyseSegmentFlop: props['ANALYSE_MCFLOP_CLIENT_1'] || "",
        
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
        
        focusKw: props['TARGET_KW'] || "",
        focusVol: props['TARGET_KW_SV'] || "",
        focusClientUrl: props['TARGET_URL_CLIENT'] || "",
        focusClientPos: props['TARGET_KW_CLIENT_POS'] || "-",
        focusNoPage: props['TARGET_KW_CLIENT_POS'] === "-" ? "true" : "false",
        focusCompUrl: props['TARGET_URL_CONCURRENT'] || "",
        focusCompPos: props['TARGET_KW_CONCURRENT_POS'] || "-",
        focusLocalisation: props['TARGET_LOCALISATION'] || "",
        
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
        intentionDescHtml: props['FOCUS_INTENTION_DESC'] || "",
        
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
        semantique3Html: props['focus_semantique_texte_3'] || "",

        // BLOC MANQUANT SLIDE 2 AJOUTÉ ICI :
        gapTitre1: props['FOCUS_GAP_TITRE_1'] || "",
        gapDesc1: props['FOCUS_GAP_DESC_1'] || "",
        gapDesc1Html: props['FOCUS_GAP_DESC_1'] || "",
        gapTitre2: props['FOCUS_GAP_TITRE_2'] || "",
        gapDesc2: props['FOCUS_GAP_DESC_2'] || "",
        gapDesc2Html: props['FOCUS_GAP_DESC_2'] || "",
        gapTitre3: props['FOCUS_GAP_TITRE_3'] || "",
        gapDesc3: props['FOCUS_GAP_DESC_3'] || "",
        gapDesc3Html: props['FOCUS_GAP_DESC_3'] || "",

        reco1: props['FOCUS_RECO_1'] || "",
        reco1Html: props['FOCUS_RECO_1'] || "",
        reco2: props['FOCUS_RECO_2'] || "",
        reco2Html: props['FOCUS_RECO_2'] || "",
        reco3: props['FOCUS_RECO_3'] || "",
        reco3Html: props['FOCUS_RECO_3'] || "",
        reco4: props['FOCUS_RECO_4'] || "",
        reco4Html: props['FOCUS_RECO_4'] || ""
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
    var urlClientIdx = -1;
    var posClientIdx = -1;
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
                for (var c = 0; c < headers.length; c++) {
                    if (String(headers[c]) === "Pos " + entName) {
                        posClientIdx = c;
                        break;
                    }
                }
            } else {
                urlCompIndices.push(j);
                for (var c = 0; c < headers.length; c++) {
                    if (String(headers[c]) === "Pos " + entName) {
                        posCompIndices.push({ posIdx: c, urlIdx: j });
                        break;
                    }
                }
            }
        }
    }
    
    Logger.log("Index URL Client : " + urlClientIdx + " | Pos Client : " + posClientIdx);
    
    var kwLower = motCle.toLowerCase();
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (String(row[0]).trim().toLowerCase() === kwLower) {
            Logger.log("Match trouvé à la ligne : " + (i + 1));
            
            var volume = row[1];
            var clientUrl = urlClientIdx > -1 ? String(row[urlClientIdx]).trim() : "";
            var clientPos = posClientIdx > -1 ? parseInt(row[posClientIdx]) : 999;
            if (isNaN(clientPos) || clientPos <= 0) clientPos = "-";
            
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
            if (bestCompPos === 9999) bestCompPos = "-";

            var result = {
                success: true,
                volume: volume,
                clientUrl: clientUrl,
                clientPos: clientPos,
                compUrl: bestCompUrl,
                compPos: bestCompPos
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

    var dicoBase64 = {
        "organique": "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",
        "paa": "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",
        "video": "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",
        "recherche": "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",
        "shopping": "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",
        "ads": "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",
        "local": "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",
        "image": "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",
        "featured": "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",
        "defaut": "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"
    };

    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var geminiApiKey = props['CONF_API_KEY_GEMINI'];
        var listeClesAPIStr = props['LISTE_CLES_API'];
        var contexteClient = props['PA_CONTEXTE_CLIENT'] || ""; 
        
        if (!geminiApiKey || geminiApiKey.trim() === "") {
            throw new Error("Clé API Gemini introuvable.");
        }
        
        var apiKeys = [];
        if (listeClesAPIStr) {
            try {
                var parsedKeys = JSON.parse(listeClesAPIStr);
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
        
        var featuresLog = serpData.elements_serp.map(function(e) { return e.type_feature; }).join(", ");
        Logger.log("SERP Features détectées : " + serpData.elements_serp.length + " (" + featuresLog + ")");
        Logger.log("Top 10 URLs récupérées : " + serpData.urls.length);

        // ==========================================
        // ÉTAPE 2 : SCRAPING EN PARALLÈLE
        // ==========================================
        Logger.log("Étape 2 : Scraping des URLs du Top 10 + URL Client");
        var urlsToScrape = serpData.urls.slice(); 
        
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
            serp_features: serpData.elements_serp.map(function(e) { return e.type_feature; }),
            pages: scrapedPages
        };

        Logger.log("Étape 3 : Création du prompt et appel API Gemini 2.5 Pro");
        
        var systemPrompt = `Tu es un expert stratège SEO. Ta mission est d'analyser des données brutes issues du scraping d'une SERP Google (top 10 concurrents + page cible du client) pour fournir un diagnostic clinique et un plan d'action stratégique sur-mesure.

            /// Protocole d'analyse (rigueur fiscale) ///

            1. Charte typographique (rigueur française absolue)
            - Titres, puces et labels : majuscule uniquement au premier mot (sauf noms propres).
            - Parenthèses : pas de majuscule au premier mot à l'intérieur.
            - Deux-points (:) : toujours un espace avant le deux-points, et aucune majuscule après (sauf nom propre).
            - Formatage visuel : utilise le gras standard Markdown (**mot**) pour mettre en valeur les concepts clés. Ne pas utiliser d'astérisque simple.

            2. Hiérarchie décisionnelle et objectifs (règle d'or : SERP > client)
            Ce que Google positionne dicte ce que tu dois recommander.
            - Mode création (si l'URL client est 'Page à créer') : ne critique pas la page client puisqu'elle n'existe pas. Projette la future page idéale basée sur les standards de la SERP.
            - Mode optimisation (si l'URL client existe) : audite la page cible existante du client par rapport aux exigences de la SERP.

            3. Intelligence commerciale et synergie (crucial)
            Ne traite pas le mot-clé cible en vase clos. Repère les expertises du client via le contexte client. Dans tes recommandations (et le gap business), propose des ponts intelligents (maillage interne, encarts de réassurance) entre le sujet analysé et l'offre globale du client.

            4. Lecture de l'environnement (SERP features)
            Examine les 'serp_features' détectées. Ce sont des signaux d'intention stricts à intégrer dans tes recommandations.

            5. Concision extrême et limites de longueur
            Tu as des limites de mots et de caractères STRICTES pour chaque champ définies dans le modèle JSON ci-dessous. Va droit au but, supprime les mots de liaison inutiles. Tu seras pénalisé si tu dépasses ces limites.

            /// Format de sortie obligatoire ///
            Tu dois fournir ta réponse uniquement sous forme d'objet JSON valide, sans balise markdown autour.
            Structure stricte exigée :
            {
            "intention": {
                "titre": "Typologie (ex: Transactionnelle) (MAX 2 mots, 65 caractères)",
                "description": "Analyse concise de l'intention. (MAX 15 mots, 70 caractères)"
            },
            "standards": [
                "Standard structurel 1. (MAX 8 mots, 65 caractères)",
                "Standard éditorial 2. (MAX 8 mots, 65 caractères)",
                "Standard UX 3. (MAX 8 mots, 65 caractères)"
            ],
            "semantique": [
                "Axe lexical 1. (MAX 8 mots, 65 caractères)",
                "Axe lexical 2. (MAX 8 mots, 65 caractères)",
                "Axe lexical 3. (MAX 8 mots, 65 caractères)"
            ],
            "gap_analysis": [
                {"titre": "Format et UX (MAX 2 mots, 25 caractères)", "description": "L'écart justifié... (MAX 25 mots, 150 caractères)"},
                {"titre": "Profondeur (MAX 2 mots, 25 caractères)", "description": "Le manque à combler... (MAX 25 mots, 150 caractères)"},
                {"titre": "Business (MAX 2 mots, 25 caractères)", "description": "L'opportunité manquée... (MAX 25 mots, 150 caractères)"}
            ],
            "recommandations": [
                "Action structure Hn/UX. (MAX 20 mots, 100 caractères)",
                "Action profondeur. (MAX 20 mots, 100 caractères)",
                "Action conversion. (MAX 20 mots, 100 caractères)",
                "Action maillage/cross-sell. (MAX 20 mots, 100 caractères)"
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

        // Ajout du Mapping SVG (Base64) pour le Front-end
        var finalElementsSerp = (serpData.elements_serp || []).map(function(el) {
            var featureKey = el.type_feature || "defaut";
            if (!dicoBase64[featureKey]) {
                featureKey = "defaut";
            }
            el.svg_icon = dicoBase64[featureKey];
            return el;
        });

        Logger.log("=== FIN : lancerWorkflowSERP (Succès) ===");
        
        // Retour Front-End
        return {
            success: true,
            data: {
                elements_serp: finalElementsSerp,
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

    var elementsSerpGeneres = [];

    // 8. Résultats organiques (placé en priorité haute)
    var organicCount = 0;
    if (json.organic_results && json.organic_results.length > 0) {
        organicCount = json.organic_results.length;
        elementsSerpGeneres.push({
            titre: organicCount + " résultats organiques",
            description: "Concurrence naturelle (1ère page).",
            type_feature: "organique"
        });
    }

    // 1. Annonces (Google Ads)
    var adsCount = 0;
    if (json.ads && json.ads.length > 0) adsCount = json.ads.length;
    else if (json.advertisements && json.advertisements.length > 0) adsCount = json.advertisements.length;
    if (adsCount > 0) {
        elementsSerpGeneres.push({
            titre: adsCount + " annonces sponsorisées",
            description: "Forte intention commerciale (enchères payantes).",
            type_feature: "ads"
        });
    }

    // 2. Position 0 (Featured Snippet)
    if (json.answer_box || json.answer_box_list) {
        elementsSerpGeneres.push({
            titre: "Position 0 (featured snippet)",
            description: "Réponse directe mise en avant par Google.",
            type_feature: "featured"
        });
    }

    // 3. Pack local (Maps)
    if ((json.local_results && (json.local_results.length > 0 || json.local_results.places)) || (json.places && json.places.length > 0)) {
        elementsSerpGeneres.push({
            titre: "Pack local (Google Maps)",
            description: "Résultats géolocalisés (zone de chalandise).",
            type_feature: "local"
        });
    }

    // 4. Bloc shopping
    if ((json.shopping_results && json.shopping_results.length > 0) || (json.inline_shopping && json.inline_shopping.length > 0) || (json.immersive_products && json.immersive_products.length > 0) || (json.commercial_units && json.commercial_units.length > 0)) {
        elementsSerpGeneres.push({
            titre: "Bloc Google Shopping",
            description: "Produits transactionnels avec prix/visuels.",
            type_feature: "shopping"
        });
    }

    // 5. Autres questions posées (PAA)
    var paaCount = 0;
    if (json.related_questions && json.related_questions.length > 0) paaCount = json.related_questions.length;
    else if (json.people_also_ask && json.people_also_ask.length > 0) paaCount = json.people_also_ask.length;
    if (paaCount > 0) {
        elementsSerpGeneres.push({
            titre: paaCount + " autres questions posées",
            description: "Questions fréquentes des internautes.",
            type_feature: "paa"
        });
    }

    // 6. Vidéos / Shorts
    if ((json.video_results && json.video_results.length > 0) || (json.inline_videos && json.inline_videos.length > 0) || (json.primetime_results && json.primetime_results.length > 0) || (json.short_videos && json.short_videos.length > 0) || (json.visual_stories && json.visual_stories.length > 0)) {
        elementsSerpGeneres.push({
            titre: "Résultats vidéos",
            description: "Intégration de formats vidéos/shorts.",
            type_feature: "video"
        });
    }

    // 7. Images
    if ((json.images_results && json.images_results.length > 0) || (json.inline_images && json.inline_images.length > 0) || (json.media_results && json.media_results.length > 0)) {
        elementsSerpGeneres.push({
            titre: "Bloc d'images",
            description: "Carrousel d'images pertinent.",
            type_feature: "image"
        });
    }

    // Limiter aux 4 premiers éléments détectés
    var finalElements = elementsSerpGeneres.slice(0, 4);

    return { urls: urls, elements_serp: finalElements };
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