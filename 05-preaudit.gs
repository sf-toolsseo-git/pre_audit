function ouvrirFenetrePreAudit() {
    var html = HtmlService.createHtmlOutputFromFile('06-preaudit')
        .setWidth(1400)
        .setHeight(1000)
        .setTitle('Pré-audit');
    SpreadsheetApp.getUi().showModelessDialog(html, '📈 Pré-audit');
}

function chargerConfigurationPreAudit() {
    var props = PropertiesService.getScriptProperties().getProperties();
    var userProps = PropertiesService.getUserProperties().getProperties();
    return {
        clientName: props['CLIENT_NAME'] || "",
        clientUrl: props['CLIENT_URL'] || "",
        urlsContexte: props['URLS_CONTEXTE'] || "",
        contexteClient: props['CONTEXTE_CLIENT'] || "",
        slideId: props['SLIDE_PRE_AUDIT_ID'] || "",
        brief: props['BRIEF_PRE_AUDIT'] || "",
        urlReponses: props['URL_REPONSES'] || "",
        contextePreaudit: props['CONTEXTE_PREAUDIT'] || "",
        besoinHtml: props['PREAUDIT_BESOIN_HTML'] || "",
        besoinTexte: props['PREAUDIT_BESOIN_TEXTE'] || "",
        solutionHtml: props['PREAUDIT_SOLUTION_HTML'] || "",
        solutionTexte: props['PREAUDIT_SOLUTION_TEXTE'] || "",
        titreSemrush: props['ANALYSE_SEMRUSH_TITRE'] || "",
        analyseKwHtml: props['ANALYSE_SEMRUSH_KW_HTML'] || "",
        analyseKwTexte: props['ANALYSE_SEMRUSH_KW'] || "",
        analyseTraficHtml: props['ANALYSE_SEMRUSH_TRAFIC_HTML'] || "",
        analyseTraficTexte: props['ANALYSE_SEMRUSH_TRAFIC'] || "",
        activeTab: userProps['PREAUDIT_ACTIVE_TAB'] || "config"
    };
}

function recupererReponseFormulaire(urlForm) {
    if (!urlForm) return "";
    
    if (urlForm.indexOf("docs.google.com/forms") === -1) {
        throw new Error("L'URL fournie n'est pas un lien Google Forms valide.");
    }

    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var clientName = (props['CLIENT_NAME'] || "").toLowerCase().trim();
        var clientUrl = (props['CLIENT_URL'] || "").toLowerCase().replace(/^(?:https?:\/\/)?(?:www\.)?/i, "").split('/')[0].trim();
        
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
    var props = PropertiesService.getScriptProperties();
    props.setProperties({
        'CLIENT_NAME': form.clientName || "",
        'CLIENT_URL': form.clientUrl || "",
        'URLS_CONTEXTE': form.urlsContexte || "",
        'CONTEXTE_CLIENT': form.contexteClient || "",
        'SLIDE_PRE_AUDIT_ID': form.slideId || "",
        'BRIEF_PRE_AUDIT': form.brief || "",
        'URL_REPONSES': form.urlReponses || "",
        'CONTEXTE_PREAUDIT': form.contextePreaudit || ""
    });
    syncPropertiesToConfigSheet();
    
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
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperty('ANALYSE_SELECTION', JSON.stringify(selection));
        return true;
    } catch (e) {
        return false;
    }
}

function chargerSelectionAnalyse() {
    try {
        var props = PropertiesService.getScriptProperties();
        var data = props.getProperty('ANALYSE_SELECTION');
        return data ? JSON.parse(data) : [];
    } catch (e) {
        return [];
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

    var clientName = props['CLIENT_NAME'] || "Client";
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

            // Segmentation SWO
            if (clientPos <= 3) {
                acquis.push({ kw: kw, vol: vol, pos: clientPos, DDT: kwDDT });
            } else if (clientPos >= 4 && clientPos <= 20) {
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
    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var apiKey = props['GEMINI_API_KEY'];
        
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
            return { success: true, jsonString: responseText };
        } else {
            throw new Error("L'API Gemini n'a renvoyé aucune analyse valide.");
        }

    } catch (e) {
        return { success: false, error: e.message };
    }
}

function sauvegarderAnalyseEvolution(titre, texteKw, texteTrafic) {
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperty('ANALYSE_SEMRUSH_TITRE', titre || "");
        props.setProperty('ANALYSE_SEMRUSH_KW', texteKw || "");
        props.setProperty('ANALYSE_SEMRUSH_TRAFIC', texteTrafic || "");
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
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperties({
            'PREAUDIT_BESOIN_HTML':          data.besoinHtml || "",
            'PREAUDIT_BESOIN_TEXTE':         data.besoinTexte || "",
            'PREAUDIT_SOLUTION_HTML':        data.solutionHtml || "",
            'PREAUDIT_SOLUTION_TEXTE':       data.solutionTexte || "",
            'ANALYSE_SEMRUSH_TITRE':         data.titreSemrush || "",
            'ANALYSE_SEMRUSH_KW_HTML':       data.analyseKwHtml || "",
            'ANALYSE_SEMRUSH_KW':            data.analyseKwTexte || "",
            'ANALYSE_SEMRUSH_TRAFIC_HTML':   data.analyseTraficHtml || "",
            'ANALYSE_SEMRUSH_TRAFIC':        data.analyseTraficTexte || ""
        });
        syncPropertiesToConfigSheet();
        return true;
    } catch (e) {
        throw new Error("Erreur lors de la sauvegarde globale : " + e.message);
    }
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
    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var apiKey = props['GEMINI_API_KEY'];
        
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
            return { success: true, texte: json.candidates[0].content.parts[0].text.trim() };
        } else {
            throw new Error("L'API Gemini n'a renvoyé aucun texte.");
        }

    } catch (error) {
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
        var apiKey = props['GEMINI_API_KEY'];
        
        if (!apiKey || apiKey.trim() === "") {
            Logger.log("Erreur : Clé API Gemini manquante.");
            throw new Error("Clé API Gemini introuvable.");
        }

        var promptStr = "Tu es un expert SEO et stratège commercial. Analyse simultanément ces deux tableaux JSON des thématiques du client : le 'Top 3' (trafic actuel) et le 'Flop 3' (manque à gagner).\n\n" +
                        "Lexique des métriques :\n" +
                        "- TEC (Trafic estimé client) : trafic actuel généré.\n" +
                        "- TPM (Trafic potentiel max) : trafic atteignable si 1ère position.\n" +
                        "- DDT (Déficit de trafic) : manque à gagner (TPM - TEC).\n\n" +
                        "Données JSON Top 3 :\n" + donneesTop + "\n\n" +
                        "Données JSON Flop 3 :\n" + donneesFlop + "\n\n" +
                        "Profilage commercial :\n" + (contexteCommercial || "Non renseigné.") + "\n\n" +
                        "RÈGLES DE RÉDACTION CRUCIALES (VERROUILLAGE SÉMANTIQUE) :\n" +
                        "1. NOMMAGE STRICT : tu dois obligatoirement utiliser le NOM EXACT des segments fournis (ex : 'Sophrologie > Général'). Il est strictement interdit de simplifier ou de généraliser (ne transforme jamais 'Sophrologie > Général' en 'La sophrologie').\n" +
                        "2. CONTEXTE '> GÉNÉRAL' : si un segment se nomme '> Général', interprète-le comme des requêtes de notoriété, de définition ou de découverte du métier (intentions informationnelles larges) et non comme l'intégralité de la thématique parente.\n" +
                        "3. STRUCTURE DES PUCES : commence chaque puce impérativement par le nom du segment concerné entre guillemets ou en gras pour ancrer l'analyse sur la donnée réelle.\n" +
                        "4. EXHAUSTIVITÉ : génère EXACTEMENT 3 puces pour le Top et EXACTEMENT 3 puces pour le Flop (une par thématique du JSON).\n" +
                        "5. INTELLIGENCE CROISÉE : si une thématique est présente dans les deux listes, souligne ce paradoxe (ex : c'est votre pilier actuel, mais aussi votre plus gros manque à gagner).\n\n" +
                        "CONTRAINTES DE STYLE ET TYPOGRAPHIE :\n" +
                        "- Style télégraphique : phrases courtes, percutantes, sans points-virgules. Un constat, un chiffre, une action.\n" +
                        "- IMPÉRATIF VISUEL : encadre les expressions clés avec des astérisques simples pour le gras orange (ex : *socle solide*, *manque à gagner colossal*).\n" +
                        "- Règles FR : un espace avant les deux-points (:), pas de majuscule après. Jours, mois et langues en minuscule. Acronymes (SEO, TEC, TPM, DDT) en majuscules.\n\n" +
                        "Format de sortie STRICTEMENT JSON :\n" +
                        "{\n" +
                        "  \"titre_slide_top\": \"Titre valorisant vos acquis\",\n" +
                        "  \"analyse_top\": [\"Le segment 'Nom exact' est un *socle solide* avec 491 visites (TEC). Notre objectif : combler le *déficit* (DDT).\", \"...\", \"...\"],\n" +
                        "  \"titre_slide_flop\": \"Titre d'alerte sur le manque à gagner\",\n" +
                        "  \"analyse_flop\": [\"Sur 'Nom exact', vous accusez un *manque à gagner colossal* (DDT de 24 965). C'est une priorité de conquête.\", \"...\", \"...\"]\n" +
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

function genererSlideBesoinSolutionIA(contextePreaudit) {
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
        var apiKey = props['GEMINI_API_KEY'];
        
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
            return { success: true, jsonString: responseText };
        } else {
            throw new Error("L'API Gemini n'a renvoyé aucune analyse valide.");
        }

    } catch (error) {
        return { success: false, message: error.message };
    }
}