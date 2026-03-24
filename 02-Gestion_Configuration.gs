function afficherFenetreConfiguration() {
    Logger.log("Ouverture de la fenêtre de configuration unifiée");
    try {
        var html = HtmlService.createHtmlOutputFromFile('03-Interface_Configuration')
            .setWidth(900)
            .setHeight(800)
            .setTitle('Configuration');
        SpreadsheetApp.getUi().showModelessDialog(html, '⚙️ Configuration');
    } catch (e) {
        Logger.log("Erreur lors de l'ouverture de la fenêtre : " + e.toString());
    }
}

function afficherFenetrePreparationClustering() {
    try {
        var html = HtmlService.createHtmlOutputFromFile('04-Interface_Preparation_Clustering')
            .setWidth(850)
            .setHeight(800)
            .setTitle('Préparation Clustering Sémantique');
        SpreadsheetApp.getUi().showModelessDialog(html, '🧩 Clustering');
    } catch (e) {
        Logger.log("Erreur affichage fenêtre clustering : " + e.toString());
    }
}

function chargerDonneesInitiales() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var oldTempSheet = ss.getSheetByName("_SAUVEGARDE_CONFIG");
    if (oldTempSheet) {
        ss.deleteSheet(oldTempSheet);
    }

    var configSheet = ss.getSheetByName("CONFIG");
    if (configSheet) {
        try {
            var data = configSheet.getDataRange().getValues();
            var propsToRestore = {};
            var hasData = false;
            
            var keyRegex = /^[A-Z0-9_]{3,}$/;
            for (var r = 1; r < data.length; r++) {
                var row = data[r];
                for (var c = 0; c < row.length; c++) {
                    var cellValue = String(row[c]).trim();
                    if (keyRegex.test(cellValue) && (c + 1) < row.length) {
                        var val = row[c + 1] !== null ? String(row[c + 1]) : "";
                        propsToRestore[cellValue] = val;
                        hasData = true;
                        c++;
                    }
                }
            }
            
            if (hasData) {
                PropertiesService.getScriptProperties().setProperties(propsToRestore);
            }
        } catch (e) {
            Logger.log("Erreur lors de la lecture de l'onglet CONFIG : " + e.toString());
        }
    }

    var props = PropertiesService.getScriptProperties().getProperties();
    var hasMatrix = (ss.getSheetByName("Matrice") !== null);
    var donnees = {
        hasMatrix: hasMatrix,
        geminiApiKey: props['GEMINI_API_KEY'] || "",
        urlsContexte: props['URLS_CONTEXTE'] || "",
        contexteClient: props['CONTEXTE_CLIENT'] || "",
        isMultiTheme: props['IS_MULTI_THEME'] === 'true',
        projectType: props['PROJECT_TYPE'] || "installe",
        clientName: props['CLIENT_NAME'] || "",
        clientUrl: props['CLIENT_URL'] || "",
        clientStrength: props['CLIENT_STRENGTH'] || "moyenne",
        clientBrand: props['CLIENT_BRAND'] || "",
        competitorName1: props['COMP_NAME_1'] || "",
        competitor1: props['COMPETITOR_1'] || "",
        competitorStrength1: props['COMP_STRENGTH_1'] || "moyenne",
        competitorBrand1: props['COMP_BRAND_1'] || "",
        competitorName2: props['COMP_NAME_2'] || "",
        competitor2: props['COMPETITOR_2'] || "",
        competitorStrength2: props['COMP_STRENGTH_2'] || "moyenne",
        competitorBrand2: props['COMP_BRAND_2'] || "",
        competitorName3: props['COMP_NAME_3'] || "",
        competitor3: props['COMPETITOR_3'] || "",
        competitorStrength3: props['COMP_STRENGTH_3'] || "moyenne",
        competitorBrand3: props['COMP_BRAND_3'] || "",
        competitorName4: props['COMP_NAME_4'] || "",
        competitor4: props['COMPETITOR_4'] || "",
        competitorStrength4: props['COMP_STRENGTH_4'] || "moyenne",
        competitorBrand4: props['COMP_BRAND_4'] || "",
        competitorName5: props['COMP_NAME_5'] || "",
        competitor5: props['COMPETITOR_5'] || "",
        competitorStrength5: props['COMP_STRENGTH_5'] || "moyenne",
        competitorBrand5: props['COMP_BRAND_5'] || ""
    };
    
    function parseAndMigrateCTR(val, defaultVal) {
        if (val === undefined || val === null || String(val).trim() === "") return defaultVal;
        var num = parseFloat(String(val).replace(',', '.'));
        if (isNaN(num)) return defaultVal;
        // Si la valeur est en décimale (ex: 0.28), on la convertit en pourcentage (28)
        // Les pourcentages habituels de CTR pour le top 10 sont généralement supérieurs à 1%
        if (num > 0 && num < 1) {
            return (num * 100).toString();
        }
        return num.toString();
    }

    donnees.ctrPos1 = parseAndMigrateCTR(props['CTR_POS_1'], "28");
    donnees.ctrPos2 = parseAndMigrateCTR(props['CTR_POS_2'], "20");
    donnees.ctrPos3 = parseAndMigrateCTR(props['CTR_POS_3'], "12");
    donnees.ctrPos4 = parseAndMigrateCTR(props['CTR_POS_4'], "8");
    donnees.ctrPos5 = parseAndMigrateCTR(props['CTR_POS_5'], "7");
    donnees.ctrPos6 = parseAndMigrateCTR(props['CTR_POS_6'], "6");
    donnees.ctrPos7 = parseAndMigrateCTR(props['CTR_POS_7'], "5");
    donnees.ctrPos8 = parseAndMigrateCTR(props['CTR_POS_8'], "5");
    donnees.ctrPos9 = parseAndMigrateCTR(props['CTR_POS_9'], "4");
    donnees.ctrPos10 = parseAndMigrateCTR(props['CTR_POS_10'], "3");

    return donnees;
}

function enregistrerConfiguration(formulaire) {
    Logger.log("Début de l'enregistrement de la configuration");
    var props = PropertiesService.getScriptProperties();
    
    // Configuration silencieuse des clés API SERP
    var listeClesAPI = {
        serpapi: ["b6f773c4a0676ee0f671b57de0d47f32ce186cf56da118af6264c6f371adefb7", "30be5fb7fab8f6d4e8a53fb1bcfc7344ffc45fe6799e981617033d344b7d7494", "5f68d13d8cc15308a0fcd5410728068a11e76faf968b5841c6c82e2fccc2a1da", "2576b1808240a807454e65cb5f8c36bd874d9c31bff711bfe9b10b53fa1c64e6", "ec4c0e232e63cc29db8728317f929348eec3d538c230c81674cd2ccb89779a01"],
        serpstack: ["3d87377fc0cf1663bf5ead3b917aea80", "82cb1d4f19dfedd7a0f4b893e05b6ad8", "abd06a2ab6bf28c2c0aa5a5b67d50416", "f765fa48d6f564b4141320993e7d2afa", "ffa21315aec4e4384702bc7556b05962"]
    };
    
    props.setProperties({
        'LISTE_CLES_API': JSON.stringify(listeClesAPI),
        'GEMINI_API_KEY': formulaire.geminiApiKey || "",
        'URLS_CONTEXTE': formulaire.urlsContexte || "",
        'CONTEXTE_CLIENT': formulaire.contexteClient || "",
        'IS_MULTI_THEME': formulaire.isMultiTheme ? "true" : "false",
        'PROJECT_TYPE': formulaire.projectType,
        'CLIENT_NAME': formulaire.clientName,
        'CLIENT_URL': formulaire.clientUrl,
        'CLIENT_STRENGTH': formulaire.clientStrength,
        'CLIENT_BRAND': formulaire.clientBrand,
        'COMP_NAME_1': formulaire.competitorName1,
        'COMPETITOR_1': formulaire.competitor1,
        'COMP_STRENGTH_1': formulaire.competitorStrength1,
        'COMP_BRAND_1': formulaire.competitorBrand1,
        'COMP_NAME_2': formulaire.competitorName2,
        'COMPETITOR_2': formulaire.competitor2,
        'COMP_STRENGTH_2': formulaire.competitorStrength2,
        'COMP_BRAND_2': formulaire.competitorBrand2,
        'COMP_NAME_3': formulaire.competitorName3,
        'COMPETITOR_3': formulaire.competitor3,
        'COMP_STRENGTH_3': formulaire.competitorStrength3,
        'COMP_BRAND_3': formulaire.competitorBrand3,
        'COMP_NAME_4': formulaire.competitorName4,
        'COMPETITOR_4': formulaire.competitor4,
        'COMP_STRENGTH_4': formulaire.competitorStrength4,
        'COMP_BRAND_4': formulaire.competitorBrand4,
        'COMP_NAME_5': formulaire.competitorName5,
        'COMPETITOR_5': formulaire.competitor5,
        'COMP_STRENGTH_5': formulaire.competitorStrength5,
        'COMP_BRAND_5': formulaire.competitorBrand5,
        'CTR_POS_1': formulaire.ctrPos1,
        'CTR_POS_2': formulaire.ctrPos2,
        'CTR_POS_3': formulaire.ctrPos3,
        'CTR_POS_4': formulaire.ctrPos4,
        'CTR_POS_5': formulaire.ctrPos5,
        'CTR_POS_6': formulaire.ctrPos6,
        'CTR_POS_7': formulaire.ctrPos7,
        'CTR_POS_8': formulaire.ctrPos8,
        'CTR_POS_9': formulaire.ctrPos9,
        'CTR_POS_10': formulaire.ctrPos10
    });
    syncPropertiesToConfigSheet();
    
    Logger.log("Configuration enregistrée avec succès");
    return { success: true };
}

function extractDomainForMatrix(url) {
    if (!url) return "";
    var match = url.match(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)/i);
    return match ? match[1].toLowerCase() : url.toLowerCase();
}

function traiterConcurrence(projectType, headersRaw, allRows) {
    try {
        if (!allRows || allRows.length === 0) throw new Error("Aucune donnée CSV trouvée.");
        
        var props = PropertiesService.getScriptProperties().getProperties();
        var isInstall = (projectType === 'installe');
        
        var clientDomain = isInstall ? extractDomainForMatrix(props['CLIENT_URL'] || "") : "";
        var clientNameRaw = props['CLIENT_NAME'] || "";
        var clientLabel = clientNameRaw.trim() !== "" ? clientNameRaw.trim() : "Client";

        var comps = [];
        for (var i = 1; i <= 5; i++) {
            var d = extractDomainForMatrix(props['COMPETITOR_' + i] || "");
            if (d) comps.push({ id: 'comp' + i, domain: d, name: props['COMP_NAME_' + i] || d });
        }
        
        if (comps.length === 0) throw new Error("Aucun concurrent configuré.");

        var kwIdx = -1, posIdx = -1, volIdx = -1, urlIdx = -1;
        for (var i = 0; i < headersRaw.length; i++) {
            var h = headersRaw[i].toLowerCase().trim();
            if (h === 'keyword' || h === 'mot-clé') kwIdx = i;
            if (h === 'position') posIdx = i;
            if (h === 'search volume' || h === 'volume') volIdx = i;
            if (h === 'url') urlIdx = i;
        }
        
        if (kwIdx === -1 || posIdx === -1 || urlIdx === -1) {
            throw new Error("Colonnes requises introuvables dans le CSV (Keyword, Position, URL).");
        }

        var mapKw = {};
        for (var i = 0; i < allRows.length; i++) {
            var r = allRows[i];
            if (!r || r.length <= Math.max(kwIdx, posIdx, urlIdx)) continue;

            var kw = String(r[kwIdx]).trim();
            if (!kw) continue;

            var pos = parseInt(r[posIdx], 10);
            var vol = parseInt(String(r[volIdx]).replace(/\s/g, ''), 10) || 0;
            var url = String(r[urlIdx]).trim();

            if (isNaN(pos) || pos <= 0 || !url) continue;

            var urlDomain = extractDomainForMatrix(url);
            if (!urlDomain) continue;

            if (!mapKw[kw]) {
                mapKw[kw] = { volume: vol, client: {pos: 999, url: ""}, competitors: {} };
                comps.forEach(function(c) { mapKw[kw].competitors[c.id] = {pos: 999, url: ""}; });
            }

            if (vol > mapKw[kw].volume) mapKw[kw].volume = vol;

            if (isInstall && clientDomain && urlDomain.indexOf(clientDomain) > -1) {
                if (pos < mapKw[kw].client.pos) { mapKw[kw].client.pos = pos; mapKw[kw].client.url = url; }
            } else {
                for (var c = 0; c < comps.length; c++) {
                    if (urlDomain.indexOf(comps[c].domain) > -1) {
                        if (pos < mapKw[kw].competitors[comps[c].id].pos) {
                            mapKw[kw].competitors[comps[c].id].pos = pos;
                            mapKw[kw].competitors[comps[c].id].url = url;
                        }
                        break;
                    }
                }
            }
        }

        var finalHeaders = ["Mot-clé", "Volume"];
        if (isInstall) finalHeaders.push("Pos " + clientLabel);
        comps.forEach(function(c) { finalHeaders.push("Pos " + c.name); });
        if (isInstall) finalHeaders.push("URL " + clientLabel);
        comps.forEach(function(c) { finalHeaders.push("URL " + c.name); });

        var finalRows = [];
        var keywords = Object.keys(mapKw);

        for (var i = 0; i < keywords.length; i++) {
            var kw = keywords[i];
            var data = mapKw[kw];
            var row = [kw, data.volume];
            
            if (isInstall) row.push(data.client.pos === 999 ? "-" : data.client.pos);
            comps.forEach(function(c) { row.push(data.competitors[c.id].pos === 999 ? "-" : data.competitors[c.id].pos); });
            if (isInstall) row.push(data.client.url || "-");
            comps.forEach(function(c) { row.push(data.competitors[c.id].url || "-"); });

            finalRows.push(row);
        }

        finalRows.sort(function(a, b) { return b[1] - a[1]; });

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var nomFeuille = "Matrice";
        
        var feuilleExistante = ss.getSheetByName(nomFeuille);
        if (feuilleExistante) ss.deleteSheet(feuilleExistante);
        
        var feuille = ss.insertSheet(nomFeuille);
        feuille.setHiddenGridlines(true);

        var fullRange = feuille.getRange(1, 1, finalRows.length + 1, finalHeaders.length);
        feuille.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
        feuille.getRange(2, 1, finalRows.length, finalHeaders.length).setValues(finalRows);
        
        fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
        
        var headerRange = feuille.getRange(1, 1, 1, finalHeaders.length);
        headerRange.setBackground("#08133B").setFontColor("white").setFontWeight("bold").setVerticalAlignment("middle").setHorizontalAlignment("left");
        
        feuille.setFrozenRows(1);
        fullRange.createFilter();

        feuille.setColumnWidth(1, 250);
        feuille.setColumnWidth(2, 80);
        feuille.getRange(2, 2, finalRows.length, 1).setNumberFormat("# ##0").setHorizontalAlignment("center");
        
        var posStartIdx = 3;
        var nbEntites = isInstall ? comps.length + 1 : comps.length;
        var urlStartIdx = posStartIdx + nbEntites;

        for (var i = 0; i < nbEntites; i++) {
            feuille.setColumnWidth(posStartIdx + i, 120);
            feuille.getRange(2, posStartIdx + i, finalRows.length, 1).setHorizontalAlignment("center");
            feuille.setColumnWidth(urlStartIdx + i, 300);
        }

        // Nettoyage des colonnes et lignes vides
        var maxCols = feuille.getMaxColumns();
        if (maxCols > finalHeaders.length) {
            feuille.deleteColumns(finalHeaders.length + 1, maxCols - finalHeaders.length);
        }
        var maxRows = feuille.getMaxRows();
        if (maxRows > finalRows.length + 1) {
            feuille.deleteRows(finalRows.length + 2, maxRows - (finalRows.length + 1));
        }

        return { status: "success", message: "Matrice générée : " + finalRows.length + " mots-clés." };

    } catch(e) {
        return { status: "error", message: e.message };
    }
}

function traiterConcurrenceFiltrer(seuilPosition, exclusionTexte, isMultiTheme) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        ss.setSpreadsheetLocale('fr_FR');
        
        var props = PropertiesService.getScriptProperties().getProperties();
        var isInstall = (props['PROJECT_TYPE'] === 'installe');
        
        var feuilleSource = ss.getSheetByName("Matrice");
        if (!feuilleSource) throw new Error("La feuille 'Matrice' est introuvable.");

        var rangeDonnees = feuilleSource.getDataRange();
        var valeurs = rangeDonnees.getValues();
        if (valeurs.length < 2) throw new Error("Pas assez de données dans la matrice.");

        var enTetes = valeurs[0];
        var nbEntites = (enTetes.length - 2) / 2;
        var posStartIdx = 2;
        var urlStartIdx = 2 + nbEntites;

        // --- GESTION DES EXCLUSIONS ---
        var rawExclusions = [];
        function addExclusion(text) {
            if (!text) return;
            var parts = text.split(',');
            parts.forEach(function(p) {
                var t = p.trim().toLowerCase();
                if (t && rawExclusions.indexOf(t) === -1) rawExclusions.push(t);
            });
        }

        addExclusion(exclusionTexte);
        addExclusion(props['CLIENT_NAME']);
        addExclusion(props['CLIENT_BRAND']);
        for (var i = 1; i <= 5; i++) {
            addExclusion(props['COMP_NAME_' + i]);
            addExclusion(props['COMP_BRAND_' + i]);
        }

        var finalExclusions = [];
        rawExclusions.forEach(function(ex) {
            finalExclusions.push(ex);
            var noAccent = supprimerAccents(ex);
            if (noAccent !== ex && finalExclusions.indexOf(noAccent) === -1) {
                finalExclusions.push(noAccent);
            }
        });

        var seuil = parseInt(seuilPosition, 10);
        if (isNaN(seuil)) seuil = 20;

        var dataRows = valeurs.slice(1);
        var lignesGardees = [];
        var minConcurrents = isMultiTheme ? 1 : 2;
        
        for (var i = 0; i < dataRows.length; i++) {
            var row = dataRows[i];
            var kw = row[0].toString().trim();
            
            // 1. Filtre d'exclusion
            var skip = false;
            var kwLC = kw.toLowerCase();
            var kwNoAccent = supprimerAccents(kwLC);
            var kwNormalized = kwNoAccent.replace(/[\s-]/g, '');

            for (var x = 0; x < finalExclusions.length; x++) {
                var brand = finalExclusions[x];
                var brandNormalized = brand.replace(/[\s-]/g, '');
                
                if (kwNoAccent.indexOf(brand) > -1) { skip = true; break; }
                if (kwNormalized.indexOf(brandNormalized) > -1) { skip = true; break; }
                
                if (brand.indexOf(' ') === -1) {
                    var words = kwNoAccent.split(/\s+/);
                    for (var w = 0; w < words.length; w++) {
                        var word = words[w];
                        if (Math.abs(word.length - brand.length) > 1) continue;
                        if (calculerLevenshtein(word, brand) <= 1) {
                            skip = true;
                            break;
                        }
                    }
                }
            }
            if (skip) continue;

            // 2. Comptage et positionnement
            var concurrentsDansLeSeuil = 0;
            var concurrentsDansTop10 = 0;
            var clientPos = 999;

            if (isInstall) {
                var cp = parseInt(row[posStartIdx], 10);
                if (!isNaN(cp) && cp > 0) clientPos = cp;
            }

            var compStartOffset = isInstall ? 1 : 0; // On ignore le client pour le comptage de la concurrence
            
            for (var c = compStartOffset; c < nbEntites; c++) {
                var pVal = parseInt(row[posStartIdx + c], 10);
                var uVal = row[urlStartIdx + c];
                
                if (!isNaN(pVal) && pVal > 0 && uVal && uVal.toString().trim() !== "-" && uVal.toString().trim() !== "") {
                    if (pVal <= seuil) concurrentsDansLeSeuil++;
                    if (pVal <= 10) concurrentsDansTop10++;
                }
            }

            // 3. Condition de conservation et Segmentation
            if (concurrentsDansLeSeuil >= minConcurrents) {
                var segment = "";
                if (isInstall && clientPos <= 10) {
                    segment = "🛡️ Acquis";
                } else if (isInstall && clientPos >= 11 && clientPos <= 20) {
                    segment = "⚡ Quick-win";
                } else {
                    var thresholdForte = isMultiTheme ? 2 : 3;
                    var thresholdPotentiel = 1;
                    
                    if (concurrentsDansTop10 >= thresholdForte) {
                        segment = "🔥 Forte concurrence";
                    } else if (concurrentsDansTop10 >= thresholdPotentiel) {
                        segment = "🎯 Potentiel validé";
                    } else {
                        segment = "💡 Opportunité (faible concurrence)";
                    }
                }
                
                // --- FILTRAGE STRICT DES CELLULES (Masquage hors seuil) ---
                var newPositions = [];
                var newUrls = [];
                
                for (var c = 0; c < nbEntites; c++) {
                    var isClient = (isInstall && c === 0);
                    var pValRaw = row[posStartIdx + c];
                    var uValRaw = row[urlStartIdx + c];
                    var pValInt = parseInt(pValRaw, 10);

                    if (isClient) {
                        // On garde toujours l'historique du client
                        newPositions.push(pValRaw);
                        newUrls.push(uValRaw);
                    } else {
                        // Pour les concurrents : on vérifie la position face au seuil
                        if (!isNaN(pValInt) && pValInt > 0 && pValInt <= seuil) {
                            newPositions.push(pValInt);
                            newUrls.push(uValRaw);
                        } else {
                            // Masquage du concurrent non pertinent sur cette ligne
                            newPositions.push("-");
                            newUrls.push("-");
                        }
                    }
                }

                var newRow = [segment, row[0], row[1]].concat(newPositions).concat(newUrls);
                lignesGardees.push(newRow);
            }
        }

        var nomFeuilleCible = "Concurrence filtrée";
        var fCible = ss.getSheetByName(nomFeuilleCible);
        if(fCible) ss.deleteSheet(fCible);
        fCible = ss.insertSheet(nomFeuilleCible);
        fCible.setHiddenGridlines(true);

        if (lignesGardees.length > 0) {
            var newHeaders = ["Segment d'attaque"].concat(enTetes);
            var fullRange = fCible.getRange(1, 1, lignesGardees.length + 1, newHeaders.length);
            
            fCible.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
            fCible.getRange(2, 1, lignesGardees.length, newHeaders.length).setValues(lignesGardees);
            fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

            var headerRange = fCible.getRange(1, 1, 1, newHeaders.length);
            headerRange.setBackground("#08133B").setFontColor("white").setFontWeight("bold").setVerticalAlignment("middle").setHorizontalAlignment("left");
        
            fCible.setFrozenRows(1);
            fullRange.createFilter();

            var backgrounds = [];
            var newPosStartIdx = 3;
            var newUrlStartIdx = 3 + nbEntites;

            for (var i = 0; i < lignesGardees.length; i++) {
                var rowFiltered = lignesGardees[i];
                var rowBackground = new Array(newHeaders.length).fill(null); 
                
                var minPos = 99999;
                var bestEntiteIndex = -1;

                // Mise en surbrillance de la meilleure URL (inclut le client si isInstall)
                for (var c = 0; c < nbEntites; c++) {
                    var valPos = parseInt(rowFiltered[newPosStartIdx + c], 10);
                    if (!isNaN(valPos) && valPos > 0 && valPos < minPos) {
                        minPos = valPos;
                        bestEntiteIndex = c;
                    }
                }

                if (bestEntiteIndex > -1) {
                    var targetUrlIndex = newUrlStartIdx + bestEntiteIndex;
                    rowBackground[targetUrlIndex] = "#d9ead3"; 
                }
                
                backgrounds.push(rowBackground);
            }
            
            fCible.getRange(2, 1, lignesGardees.length, newHeaders.length).setBackgrounds(backgrounds);

            fCible.setColumnWidth(1, 200); // Segment
            fCible.setColumnWidth(2, 250); // Mot-clé
            fCible.setColumnWidth(3, 80); // Volume
            fCible.getRange(2, 3, lignesGardees.length, 1).setNumberFormat("# ### ##0").setHorizontalAlignment("center");

            for (var i = 0; i < nbEntites; i++) {
                fCible.setColumnWidth(newPosStartIdx + i, 120);
                fCible.getRange(2, newPosStartIdx + i, lignesGardees.length, 1).setHorizontalAlignment("center");
                fCible.setColumnWidth(newUrlStartIdx + i, 300);
            }

            // Nettoyage des colonnes et lignes vides
            var maxCols = fCible.getMaxColumns();
            if (maxCols > newHeaders.length) {
                fCible.deleteColumns(newHeaders.length + 1, maxCols - newHeaders.length);
            }
            var maxRows = fCible.getMaxRows();
            if (maxRows > lignesGardees.length + 1) {
                fCible.deleteRows(lignesGardees.length + 2, maxRows - (lignesGardees.length + 1));
            }
        }

        return { status: "success", message: "Filtrage terminé : " + lignesGardees.length + " mots-clés conservés." };

    } catch (e) {
        return { status: "error", message: e.message };
    }
}

function supprimerAccents(texte) {
    if (!texte) return "";
    return texte.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function calculerLevenshtein(a, b) {
    if (a.length === 0) return b.length;
    if (b.length === 0) return a.length;
    var matrix = [];
    for (var i = 0; i <= b.length; i++) { matrix[i] = [i]; }
    for (var j = 0; j <= a.length; j++) { matrix[0][j] = j; }
    for (var i = 1; i <= b.length; i++) {
        for (var j = 1; j <= a.length; j++) {
            if (b.charAt(i - 1) === a.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1,
                    Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1)
                );
            }
        }
    }
    return matrix[b.length][a.length];
}

function recupererDonneesBrutesClustering(contexteClient) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var props = PropertiesService.getScriptProperties().getProperties();
        
        // Récupération correcte du paramètre conservé en base (coché ou non)
        var isMultiTheme = (props['IS_MULTI_THEME'] === 'true');
        
        // Priorité au contexte envoyé depuis l'interface, sinon on prend celui en base
        var ctxFinal = contexteClient || props['CONTEXTE_CLIENT'] || "";

        var sheet = ss.getSheetByName("Concurrence filtrée");
        if (!sheet) throw new Error("L'onglet 'Concurrence filtrée' est introuvable. Veuillez d'abord générer la vue filtrée.");
        
        var data = sheet.getDataRange().getValues();
        if (data.length < 2) throw new Error("Aucune donnée à exporter dans 'Concurrence filtrée'.");
        
        var headers = data[0];
        // Structure de "Concurrence filtrée" : Segment(0), KW(1), Vol(2), Pos...(3 à 3+nb-1), URLs...(3+nb à fin)
        var nbEntites = (headers.length - 3) / 2;
        var urlStartIdx = 3 + nbEntites;

        var exportList = [];
        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var kw = String(row[1]).trim();
            var vol = parseInt(row[2], 10) || 0;
            
            if (kw === "") continue;

            var urlsFound = [];
            for (var c = 0; c < nbEntites; c++) {
                var pos = parseInt(row[3 + c], 10);
                var url = String(row[urlStartIdx + c]);

                // On ne récupère l'URL que si le concurrent est dans le top 20
                if (!isNaN(pos) && pos > 0 && pos <= 20 && url && url !== "-" && url.trim() !== "") {
                    var cleanUrl = url.trim();
                    if (urlsFound.indexOf(cleanUrl) === -1) {
                        urlsFound.push(cleanUrl);
                    }
                }
            }

            exportList.push({
                "keyword": kw,
                "volume": vol,
                "urls": urlsFound
            });
        }

        return {
            source: "export_seo_multi_urls",
            count: exportList.length,
            mode_multi_thematique: isMultiTheme,
            contexte_client: ctxFinal,
            keywords: exportList
        };
    } catch (e) {
        return { error: e.message };
    }
}

function preparerDonneesClustering(jsonMotsCles) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var props = PropertiesService.getScriptProperties().getProperties();
        var isInstall = (props['PROJECT_TYPE'] === 'installe');

        var sheetSource = ss.getSheetByName("Concurrence filtrée");
        var mapKw = {};

        // 1. Indexation de la matrice "Concurrence filtrée" pour réconciliation
        if (sheetSource) {
            var data = sheetSource.getDataRange().getValues();
            if (data.length > 1) {
                var headers = data[0];
                var nbEntites = (headers.length - 3) / 2;
                var posStartIdx = 3;
                var urlStartIdx = 3 + nbEntites;
                var compStartOffset = isInstall ? 1 : 0;

                for (var i = 1; i < data.length; i++) {
                    var row = data[i];
                    var kw = String(row[1]).trim().toLowerCase();

                    var bestCompPos = 99999;
                    var bestCompUrl = "";
                    var clientUrl = "-";
                    var clientPos = "-";

                    // Extraction Client
                    if (isInstall) {
                        var cPos = parseInt(row[posStartIdx], 10);
                        var cUrl = row[urlStartIdx];
                        
                        if (!isNaN(cPos) && cPos > 0) {
                            clientPos = cPos;
                        }
                        if (cUrl && String(cUrl).trim() !== "-") {
                            clientUrl = cUrl;
                        }
                    }

                    // Extraction Concurrents (URL Master)
                    for (var c = compStartOffset; c < nbEntites; c++) {
                        var pVal = parseInt(row[posStartIdx + c], 10);
                        var uVal = row[urlStartIdx + c];
                        if (!isNaN(pVal) && pVal > 0 && pVal < bestCompPos && uVal && String(uVal).trim() !== "-") {
                            bestCompPos = pVal;
                            bestCompUrl = uVal;
                        }
                    }

                    mapKw[kw] = {
                        bestCompPos: bestCompPos === 99999 ? "-" : bestCompPos,
                        bestCompUrl: bestCompUrl || "-",
                        clientUrl: clientUrl,
                        clientPos: clientPos
                    };
                }
            }
        }

        // 2. Traitement du JSON
        var clusters = JSON.parse(jsonMotsCles);
        var rowsFinal = [];

        for (var i = 0; i < clusters.length; i++) {
            var item = clusters[i];
            var kwMain = String(item.keyword).trim();
            var kwLookup = kwMain.toLowerCase();
            var univers = item.univers || "Général";
            var subTheme = item.sub_theme || "Général";
            var intent = item.intent || "-";
            var vol = parseInt(item.volume, 10) || 0;
            var variants = Array.isArray(item.variants) ? item.variants : [];

            var mappedData = mapKw[kwLookup] || { bestCompPos: "-", bestCompUrl: "-", clientUrl: "-", clientPos: "-" };
            var variantsStr = variants.join(", ");
            var nbKw = 1 + variants.length;

            // NOUVEL ORDRE DES COLONNES
            var rowData = [
                univers,
                subTheme,
                nbKw,
                kwMain,
                vol,
                intent,
                mappedData.bestCompUrl,
                mappedData.bestCompPos
            ];

            // Insertion dynamique de l'URL Client et Position Client
            if (isInstall) {
                rowData.push(mappedData.clientUrl);
                rowData.push(mappedData.clientPos);
            }

            rowData.push(variantsStr);
            rowsFinal.push(rowData);
        }

        // 3. Tri Intelligent (A-Z > A-Z > Vol Desc | Hors périmètre à la fin)
        rowsFinal.sort(function(a, b) {
            var uA = String(a[0]);
            var uB = String(b[0]);
            var isAOut = (uA.toLowerCase() === "hors périmètre");
            var isBOut = (uB.toLowerCase() === "hors périmètre");

            if (isAOut && !isBOut) return 1;
            if (!isAOut && isBOut) return -1;

            var cmpU = uA.localeCompare(uB);
            if (cmpU !== 0) return cmpU;

            var sA = String(a[1]);
            var sB = String(b[1]);
            var cmpS = sA.localeCompare(sB);
            if (cmpS !== 0) return cmpS;

            return b[4] - a[4]; // Tri par volume (Index 4)
        });

        // 4. Création de l'onglet "Cluster"
        var nomFeuilleCible = "Cluster";
        var sheetCible = ss.getSheetByName(nomFeuilleCible);
        if (sheetCible) {
            ss.deleteSheet(sheetCible);
        }
        sheetCible = ss.insertSheet(nomFeuilleCible);
        sheetCible.setHiddenGridlines(true);

        var finalHeaders = [
            "Thématique", "Sous-thématique", "Nb mots-clés", "Mot-clé principal",
            "Volume", "Intention", "URL Master", "Position master"
        ];
        
        if (isInstall) {
            finalHeaders.push("URL Client", "Position client");
        }
        
        finalHeaders.push("Mots-clés secondaires");

        if (rowsFinal.length > 0) {
            var fullRange = sheetCible.getRange(1, 1, rowsFinal.length + 1, finalHeaders.length);
            sheetCible.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
            sheetCible.getRange(2, 1, rowsFinal.length, finalHeaders.length).setValues(rowsFinal);

            fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

            var headerRange = sheetCible.getRange(1, 1, 1, finalHeaders.length);
            headerRange.setBackground("#08133B").setFontColor("white").setFontWeight("bold").setVerticalAlignment("middle").setHorizontalAlignment("left");

            sheetCible.setFrozenRows(1);
            fullRange.createFilter();

            // 1. Alignement vertical au centre pour tout l'onglet
            fullRange.setVerticalAlignment("middle");

            // 2. Formatage des largeurs de colonnes
            sheetCible.setColumnWidth(1, 200); // Thématique
            sheetCible.setColumnWidth(2, 350); // Sous-Thème
            sheetCible.setColumnWidth(3, 100); // Nb mots-clés
            sheetCible.setColumnWidth(4, 300); // Mot-clé principal
            sheetCible.setColumnWidth(5, 100); // Volume
            sheetCible.setColumnWidth(6, 150); // Intention
            sheetCible.setColumnWidth(7, 750); // URL Master
            sheetCible.setColumnWidth(8, 120); // Position master
            
            var currentIdx = 9;
            if (isInstall) {
                sheetCible.setColumnWidth(currentIdx, 750); // URL Client
                currentIdx++;
                sheetCible.setColumnWidth(currentIdx, 120); // Position client
                currentIdx++;
            }
            
            sheetCible.setColumnWidth(currentIdx, 500); // Mots-clés secondaires

            // 3. Alignements horizontaux et formatages spécifiques (données uniquement)
            sheetCible.getRange(2, 3, rowsFinal.length, 1).setHorizontalAlignment("center"); // Nb mots-clés
            sheetCible.getRange(2, 5, rowsFinal.length, 1).setNumberFormat("# ### ##0").setHorizontalAlignment("center"); // Volume
            sheetCible.getRange(2, 6, rowsFinal.length, 1).setHorizontalAlignment("center"); // Intention
            sheetCible.getRange(2, 8, rowsFinal.length, 1).setHorizontalAlignment("center"); // Position master
            
            if (isInstall) {
                sheetCible.getRange(2, currentIdx - 1, rowsFinal.length, 1).setHorizontalAlignment("center"); // Position client
            }
            
            // Wrap text pour les mots-clés secondaires
            sheetCible.getRange(2, currentIdx, rowsFinal.length, 1).setWrap(true);

            // Nettoyage visuel
            var maxCols = sheetCible.getMaxColumns();
            if (maxCols > finalHeaders.length) {
                sheetCible.deleteColumns(finalHeaders.length + 1, maxCols - finalHeaders.length);
            }
            var maxRows = sheetCible.getMaxRows();
            if (maxRows > rowsFinal.length + 1) {
                sheetCible.deleteRows(rowsFinal.length + 2, maxRows - (rowsFinal.length + 1));
            }
        }

        return { status: "success", message: "Clusters importés avec succès : " + rowsFinal.length + " grappes formées." };

    } catch (e) {
        return { status: "error", message: e.message };
    }
}

function genererContexteClientIA(urlsTexte) {
    try {
        var props = PropertiesService.getScriptProperties().getProperties();
        var apiKey = props['GEMINI_API_KEY'];
        
        if (!apiKey || apiKey.trim() === "") {
            throw new Error("Clé API Gemini introuvable. Veuillez configurer et sauvegarder l'onglet Général.");
        }

        var clientUrl = props['CLIENT_URL'] || "";
        var urlsArray = urlsTexte ? urlsTexte.split('\n') : [];
        
        if (clientUrl && clientUrl.trim() !== "") {
            urlsArray.unshift(clientUrl);
        }

        var urlsPropres = [];
        for (var i = 0; i < urlsArray.length; i++) {
            var urlStr = urlsArray[i].trim();
            if (urlStr !== "" && urlsPropres.indexOf(urlStr) === -1) {
                urlsPropres.push(urlStr);
            }
        }

        if (urlsPropres.length === 0) {
            throw new Error("Aucune URL valide à analyser.");
        }

        var contenuGlobal = "";
        for (var j = 0; j < urlsPropres.length; j++) {
            try {
                var response = UrlFetchApp.fetch(urlsPropres[j], { muteHttpExceptions: true, timeout: 10000 });
                if (response.getResponseCode() === 200) {
                    var html = response.getContentText();
                    var $ = Cheerio.load(html);
                    
                    $('script, style, nav, footer, header, aside, noscript, svg, form, iframe').remove();
                    
                    var textePage = $('p, h1, h2, h3, h4, h5, h6, li').map(function() {
                        return $(this).text().trim();
                    }).get().join('\n');
                    
                    contenuGlobal += "--- Contenu de " + urlsPropres[j] + " ---\n" + textePage + "\n\n";
                }
            } catch (e) {
                Logger.log("Erreur scraping " + urlsPropres[j] + " : " + e.message);
            }
        }

        if (!contenuGlobal.trim()) {
            throw new Error("Extraction impossible. Les sites sont vides ou protégés contre le scraping.");
        }

        if (contenuGlobal.length > 25000) {
            contenuGlobal = contenuGlobal.substring(0, 25000);
        }

        var promptStr = "Tu es un expert SEO. À partir du contenu brut extrait du site web du client, tu dois générer son \"contexte métier\" selon la structure stricte ci-dessous.\n" +
                        "Extrais un maximum d'informations pertinentes pour nourrir la compréhension globale d'une IA, sans inventer de données.\n\n" +
                        "RÈGLES TYPOGRAPHIQUES OBLIGATOIRES (français) :\n" +
                        "1. Majuscule uniquement au premier mot des labels (sauf noms propres).\n" +
                        "2. Pas de majuscule au premier mot à l'intérieur d'une parenthèse (sauf nom propre).\n" +
                        "3. Jours, mois et langues toujours en minuscule.\n" +
                        "4. Pas de majuscule après les deux points (:) car ce ne sont pas des phrases complètes.\n\n" +
                        "STRUCTURE ATTENDUE :\n" +
                        "- Modèle économique : (ex : e-commerce b2b, génération de leads locaux, saas...)\n" +
                        "- Secteur d'activité principal : (définition précise du cœur de métier)\n" +
                        "- Proposition de valeur unique : (ce qui différencie l'offre sur le marché)\n" +
                        "- Positionnement et gamme : (ex : premium, accessible, sur-mesure...)\n" +
                        "- Cadre réglementaire et certifications : (ex : qualiopi, cpf, diplôme d'état...)\n" +
                        "- Périmètre géographique : (ex : national fr, local ciblé, international...)\n" +
                        "- Typologie d'audience : (ex : b2b cible direction, b2c grand public...)\n" +
                        "- Définition de la conversion cible : (ex : achat en ligne, prise de rdv, demande de devis...)\n\n" +
                        "Texte brut du site :\n" + contenuGlobal;

        var payload = {
            "contents": [{
                "parts": [{"text": promptStr}]
            }]
        };

        var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-pro-preview:generateContent";
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

function chargerContexteIA() {
    var props = PropertiesService.getScriptProperties().getProperties();
    return {
        urlsContexte: props['URLS_CONTEXTE'] || "",
        contexteClient: props['CONTEXTE_CLIENT'] || ""
    };
}

function sauvegarderContexteIA(urls, contexte) {
    PropertiesService.getScriptProperties().setProperties({
        'URLS_CONTEXTE': urls || "",
        'CONTEXTE_CLIENT': contexte || ""
    });
    
    // On synchronise vers l'onglet
    syncPropertiesToConfigSheet();
    
    return true;
}