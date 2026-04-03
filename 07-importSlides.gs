function appliquerMarkdownSurForme(element) {
    try {
        var textRange = element.getText();
        var textStr = textRange.asString();
        
        var regex = /\*\*([^*]+)\*\*|\*([^*]+)\*/g;
        var matches = [];
        var match;
        
        while ((match = regex.exec(textStr)) !== null) {
            if (match[1]) {
                // **texte**
                matches.push({
                    start: match.index,
                    textLength: match[1].length,
                    type: 'bold'
                });
            } else if (match[2]) {
                // *texte*
                matches.push({
                    start: match.index,
                    textLength: match[2].length,
                    type: 'orange'
                });
            }
        }
        
        for (var i = matches.length - 1; i >= 0; i--) {
            var m = matches[i];
            
            if (m.type === 'bold') {
                var endMarker = m.start + m.textLength + 2;
                textRange.getRange(endMarker, endMarker + 2).clear();
                textRange.getRange(m.start, m.start + 2).clear();
                
                var styledRange = textRange.getRange(m.start, m.start + m.textLength);
                styledRange.getTextStyle().setBold(true);
            } else if (m.type === 'orange') {
                var endMarker = m.start + m.textLength + 1;
                textRange.getRange(endMarker, endMarker + 1).clear();
                textRange.getRange(m.start, m.start + 1).clear();
                
                var styledRange = textRange.getRange(m.start, m.start + m.textLength);
                styledRange.getTextStyle().setBold(true).setForegroundColor("#f67604");
            }
        }
    } catch(e) {
        // Ignorer silencieusement si la forme n'a pas de texte ou si getText() n'est pas supporté
    }
}

function exporterSlideBesoinSolution(texteBesoin, texteSolution) {
    try {
        Logger.log("=== DÉBUT : exporterSlideBesoinSolution ===");
        var props = getDatabaseData();
        var slideId = props['PA_CONF_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();
        
        var tagsTrouves = 0;
        
        // Formatage : on ajoute un saut de ligne vide entre chaque puce pour aérer dans Slides
        var slideTexteBesoin = texteBesoin.replace(/\n/g, '\n\n');
        var slideTexteSolution = texteSolution.replace(/\n/g, '\n\n');
        
        Logger.log("Recherche des tags PA_GLOBALE_BESOIN et PA_GLOBALE_SOLUTION via description");

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();
            
            shapes.forEach(function(shape) {
                var descRaw = shape.getDescription() || "";
                var targetKey = null;

                // Détection via le texte alternatif en majuscule UNIQUEMENT (Cas 1)
                if (descRaw === "PA_GLOBALE_BESOIN") {
                    targetKey = "besoin";
                } else if (descRaw === "PA_GLOBALE_SOLUTION") {
                    targetKey = "solution";
                }

                if (targetKey === "besoin") {
                    Logger.log("Forme cible 'besoin' trouvée, écrasement du texte");
                    shape.getText().setText(slideTexteBesoin);
                    appliquerMarkdownSurForme(shape);
                    tagsTrouves++;
                } else if (targetKey === "solution") {
                    Logger.log("Forme cible 'solution' trouvée, écrasement du texte");
                    shape.getText().setText(slideTexteSolution);
                    appliquerMarkdownSurForme(shape);
                    tagsTrouves++;
                }
            });
        });
        
        Logger.log("Tags trouvés et remplacés : " + tagsTrouves);
        Logger.log("=== FIN : exporterSlideBesoinSolution ===");
        
        if (tagsTrouves === 0) {
            return { success: false, error: "Les tags 'PA_GLOBALE_BESOIN' et 'PA_GLOBALE_SOLUTION' n'ont pas été trouvés dans le texte alternatif de la présentation." };
        }

        return { success: true, url: presentation.getUrl() };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT BESOIN/SOLUTION : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterAnalyseSemrushSlide(titre, texteKw, texteTrafic, imgKwB64, imgKwMime, imgTraficB64, imgTraficMime) {
    try {
        Logger.log("=== DÉBUT : exporterAnalyseSemrushSlide ===");
        var props = getDatabaseData();
        var slideId = props['PA_CONF_SLIDE_ID'];
        
        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        Logger.log("Parcours des slides pour l'analyse Semrush en cours");

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();
            
            shapes.forEach(function(shape) {
                var descRaw = shape.getDescription() || "";

                // Cas 1 : Description correspondante -> Écrasement total
                if (descRaw === "PA_GLOBALE_TITRE_SEMRUSH") {
                    Logger.log("Remplacement du titre PA_GLOBALE_TITRE_SEMRUSH");
                    shape.getText().setText(titre);
                }
                
                if (descRaw === "PA_GLOBALE_SEMRUSH_MOTCLE") {
                    Logger.log("Remplacement et formatage PA_GLOBALE_SEMRUSH_MOTCLE");
                    shape.getText().setText(texteKw);
                    appliquerMarkdownSurForme(shape);
                }
                
                if (descRaw === "PA_GLOBALE_SEMRUSH_TRAFIC") {
                    Logger.log("Remplacement et formatage PA_GLOBALE_SEMRUSH_TRAFIC");
                    shape.getText().setText(texteTrafic);
                    appliquerMarkdownSurForme(shape);
                }

                // Cas 1 sur Image (Placeholders)
                if (descRaw === "PA_GLOBALE_PLACEHOLDER_SEMRUSH_MOTCLE") {
                    Logger.log("Remplacement image mots-clés");
                    var blobKw = Utilities.newBlob(Utilities.base64Decode(imgKwB64), imgKwMime, "kw.png");
                    var newImageKw = slide.insertImage(blobKw, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                    newImageKw.setDescription(descRaw);
                    shape.remove();
                }
                
                if (descRaw === "PA_GLOBALE_PLACEHOLDER_SEMRUSH_TRAFIC") {
                    Logger.log("Remplacement image trafic");
                    var blobTrafic = Utilities.newBlob(Utilities.base64Decode(imgTraficB64), imgTraficMime, "trafic.png");
                    var newImageTrafic = slide.insertImage(blobTrafic, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                    newImageTrafic.setDescription(descRaw);
                    shape.remove();
                }
            });
        });
        Logger.log("=== FIN : exporterAnalyseSemrushSlide ===");
        return { success: true, url: presentation.getUrl() };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT SLIDE SEMRUSH : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterPerformanceGlobalSlides(diagnosticData, iaData, concurrenceData) {
    try {
        Logger.log("=== DÉBUT : exporterPerformanceGlobalSlides ===");
        var props = getDatabaseData();
        var slideId = props['PA_CONF_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        var clientKpi = diagnosticData.kpis.find(function(k) { return k.isClient; });
        if (!clientKpi) throw new Error("Données client introuvables dans le diagnostic.");

        var intentStats = diagnosticData.intentStats;
        var totalTop10 = clientKpi.top10;
        var transacPctDec = totalTop10 > 0 ? (intentStats.transac.top10 / totalTop10) : 0;
        var infoPctDec = totalTop10 > 0 ? (intentStats.info.top10 / totalTop10) : 0;
        
        Logger.log("Préparation des données pour injection...");
        var topThemes = diagnosticData.themeStats ? diagnosticData.themeStats.slice().sort(function(a, b) { return b.TEC - a.TEC || b.top10 - a.top10 || b.top3 - a.top3; }).slice(0, 3) : [];
        var flopThemes = diagnosticData.themeStats ? diagnosticData.themeStats.slice().sort(function(a, b) { return b.DDT - a.DDT; }).slice(0, 3) : [];
        var acquis = diagnosticData.acquis ? diagnosticData.acquis.slice(0, 5) : [];
        var gains = diagnosticData.gains ? diagnosticData.gains.slice(0, 5) : [];
        var pertes = diagnosticData.pertes ? diagnosticData.pertes.slice(0, 5) : [];
        var territoires = diagnosticData.territoires ? diagnosticData.territoires.slice(0, 5) : [];
        
        function safeNum(val) {
            return (val !== null && val !== undefined && !isNaN(val)) ? Math.round(val).toLocaleString('fr-FR') : "-";
        }
        function safePos(val) {
            return (val !== null && val !== undefined && !isNaN(val)) ? Number(val).toLocaleString('fr-FR') : "-";
        }

        var mappingComp = {};
        if (concurrenceData) {
            mappingComp['PA_ETAT_TITRE_CONCURRENCE'] = "L'environnement concurrentiel de " + (concurrenceData.client ? concurrenceData.client.name : "");
            if (concurrenceData.client) {
                mappingComp['PA_ETAT_NOM_CLIENT'] = concurrenceData.client.name;
                mappingComp['PA_ETAT_TOP10_CLIENT'] = safeNum(concurrenceData.client.top10);
                mappingComp['PA_ETAT_PAGES_CLIENT'] = safeNum(concurrenceData.client.pages);
            }
            if (concurrenceData.leader) {
                mappingComp['PA_ETAT_NOM_LEADER'] = concurrenceData.leader.name;
                mappingComp['PA_ETAT_TOP10_LEADER'] = safeNum(concurrenceData.leader.top10);
                mappingComp['PA_ETAT_PAGES_LEADER'] = safeNum(concurrenceData.leader.pages);
            }
            for (var c = 1; c <= 4; c++) {
                var comp = concurrenceData.comps && concurrenceData.comps[c-1] ? concurrenceData.comps[c-1] : null;
                if (comp) {
                    mappingComp['PA_ETAT_NOM_COMP' + c] = comp.name;
                    mappingComp['PA_ETAT_TOP10_COMP' + c] = safeNum(comp.top10);
                    mappingComp['PA_ETAT_PAGES_COMP' + c] = safeNum(comp.pages);
                }
            }
        }

        var mapping = {
            'PA_ETAT_MOTCLE_CLIENT_GLOBAL': (clientKpi.posAll || 0).toLocaleString('fr-FR'),
            'PA_ETAT_MOTCLE_CLIENT_TOP3': (clientKpi.top3 || 0).toLocaleString('fr-FR'),
            'PA_ETAT_MOTCLE_CLIENT_TOP10': (clientKpi.top10 || 0).toLocaleString('fr-FR'),
            'PA_ETAT_CLIENT_URL': (clientKpi.urlsCount || 0).toLocaleString('fr-FR'),
            'PA_ETAT_MOTCLE_CLIENT_TRANSAC': (intentStats.transac.top100 || 0).toLocaleString('fr-FR'),
            'PA_ETAT_MOTCLE_CLIENT_INFO': (intentStats.info.top100 || 0).toLocaleString('fr-FR'),
            'PA_ETAT_MOTCLE_CLIENT_TRANSAC_TOP10': (intentStats.transac.top10 || 0).toLocaleString('fr-FR'),
            'PA_ETAT_MOTCLE_CLIENT_INFO_TOP10': (intentStats.info.top10 || 0).toLocaleString('fr-FR'),
            'PA_ETAT_MOTCLE_CLIENT_TRANSAC_PCT': Math.round(transacPctDec * 100) + "%",
            'PA_ETAT_MOTCLE_CLIENT_INFO_PCT': Math.round(infoPctDec * 100) + "%"
        };
        
        var replaceDict = {};
        for (var i = 1; i <= 3; i++) {
            var thm = topThemes[i - 1];
            replaceDict["{{top_thm_client_" + i + "}}"] = thm ? thm.name : "-";
            replaceDict["{{top_thm_client_top10_" + i + "}}"] = thm ? safeNum(thm.top10) : "-";
            replaceDict["{{top_thm_client_tec_" + i + "}}"] = thm ? safeNum(thm.TEC) : "-";
            replaceDict["{{top_thm_client_tpm_" + i + "}}"] = thm ? safeNum(thm.TPM) : "-";
            replaceDict["{{top_thm_client_ddt_" + i + "}}"] = thm ? safeNum(thm.DDT) : "-";
        }

        for (var i = 1; i <= 3; i++) {
            var thm = flopThemes[i - 1];
            replaceDict["{{flop_thm_client_" + i + "}}"] = thm ? thm.name : "-";
            replaceDict["{{flop_thm_client_flop10_" + i + "}}"] = thm ? safeNum(thm.top10) : "-";
            replaceDict["{{flop_thm_client_tec_" + i + "}}"] = thm ? safeNum(thm.TEC) : "-";
            replaceDict["{{flop_thm_client_tpm_" + i + "}}"] = thm ? safeNum(thm.TPM) : "-";
            replaceDict["{{flop_thm_client_ddt_" + i + "}}"] = thm ? safeNum(thm.DDT) : "-";
        }

        for (var i = 1; i <= 5; i++) {
            var acq = acquis[i - 1];
            replaceDict["{{top_mc_client_" + i + "}}"] = acq ? acq.kw : "-";
            replaceDict["{{top_mc_client_vol_" + i + "}}"] = acq ? safeNum(acq.vol) : "-";
            replaceDict["{{top_mc_client_ddt_" + i + "}}"] = acq ? safeNum(acq.DDT) : "-";
            replaceDict["{{top_mc_client_pos_" + i + "}}"] = acq ? safePos(acq.pos) : "-";

            var gn = gains[i - 1];
            replaceDict["{{qw_mc_client_" + i + "}}"] = gn ? gn.kw : "-";
            replaceDict["{{qw_mc_client_vol_" + i + "}}"] = gn ? safeNum(gn.vol) : "-";
            replaceDict["{{qw_mc_client_ddt_" + i + "}}"] = gn ? safeNum(gn.DDT) : "-";
            replaceDict["{{qw_mc_client_pos_" + i + "}}"] = gn ? safePos(gn.pos) : "-";

            var prt = pertes[i - 1];
            replaceDict["{{pc_mc_client_" + i + "}}"] = prt ? prt.kw : "-";
            replaceDict["{{pc_mc_client_vol_" + i + "}}"] = prt ? safeNum(prt.vol) : "-";
            replaceDict["{{pc_mc_client_ddt_" + i + "}}"] = prt ? safeNum(prt.DDT) : "-";
            replaceDict["{{pc_mc_conc_pos_" + i + "}}"] = prt ? safePos(prt.bestCompPos) : "-";

            var terr = territoires[i - 1];
            replaceDict["{{tap_mc_client_" + i + "}}"] = terr ? terr.kw : "-";
            replaceDict["{{tap_mc_client_vol_" + i + "}}"] = terr ? safeNum(terr.vol) : "-";
            replaceDict["{{tap_mc_client_ddt_" + i + "}}"] = terr ? safeNum(terr.DDT) : "-";
            replaceDict["{{tap_mc_conc_pos_" + i + "}}"] = terr ? safePos(terr.bestPos) : "-";
        }

        function splitAnalysis(text) {
            if (!text) return ["", "", ""];
            var parts = text.split(/(?:^|\n)[•-]\s*/).map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });
            if (parts.length === 1 && text.indexOf('\n\n') !== -1) {
                parts = text.split('\n\n').map(function(s) { return s.trim().replace(/^[•-]\s*/, ''); });
            }
            return [parts[0] || "", parts[1] || "", parts[2] || ""];
        }

        var topThemParts = splitAnalysis(iaData ? iaData.analyseTopThematiques : "");
        var flopThemParts = splitAnalysis(iaData ? iaData.analyseFlopThematiques : "");
        var topSegParts = splitAnalysis(iaData ? iaData.analyseTopSegments : "");
        var flopSegParts = splitAnalysis(iaData ? iaData.analyseFlopSegments : "");

        var propsToSave = {};
        for (var k in mapping) { propsToSave[k] = String(mapping[k]); }
        for (var k in mappingComp) { propsToSave[k] = String(mappingComp[k]); }
        for (var k in replaceDict) {
            var cleanKey = k.replace(/[{}]/g, '');
            propsToSave[cleanKey] = String(replaceDict[k]);
        }
        for (var idx = 1; idx <= 3; idx++) {
            propsToSave["PA_ETAT_ANALYSE_THEMATIQUETOP" + idx] = topThemParts[idx-1];
            propsToSave["PA_ETAT_ANALYSE_THEMATIQUEFLOP" + idx] = flopThemParts[idx-1];
            propsToSave["PA_ETAT_ANALYSE_MCTOP" + idx] = topSegParts[idx-1];
            propsToSave["PA_ETAT_ANALYSE_MCFLOP" + idx] = flopSegParts[idx-1];
        }

        propsToSave["PLACEHOLDER_LOGO_CLIENT"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_LEADER"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP1"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP2"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP3"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP4"] = "IMAGE";

        setDatabaseData(propsToSave);

        Logger.log("Parcours minutieux des slides pour Performance Globale");

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();

            shapes.forEach(function(shape) {
                var shapeText = "";
                try {
                    shapeText = shape.getText().asString();
                } catch(e) {}
                var descRaw = shape.getDescription() || "";

                // --- Cas 1 : Remplacement direct via description exacte ---
                if (mapping[descRaw] !== undefined) {
                    shape.getText().setText(mapping[descRaw].toString());
                }
                if (mappingComp[descRaw] !== undefined) {
                    shape.getText().setText(mappingComp[descRaw].toString());
                }

                // Titres IA
                if (descRaw === "PA_ETAT_TITRE_THEMATIQUETOP" && iaData && iaData.titreTopThematiques) {
                    shape.getText().setText(iaData.titreTopThematiques);
                }
                if (descRaw === "PA_ETAT_TITRE_THEMATIQUEFLOP" && iaData && iaData.titreFlopThematiques) {
                    shape.getText().setText(iaData.titreFlopThematiques);
                }
                if (descRaw === "PA_ETAT_TITRE_MCTOP" && iaData && iaData.titreTopSegments) {
                    shape.getText().setText(iaData.titreTopSegments);
                }
                if (descRaw === "PA_ETAT_TITRE_MCFLOP" && iaData && iaData.titreFlopSegments) {
                    shape.getText().setText(iaData.titreFlopSegments);
                }

                // Analyses IA découpées
                for (var idx = 1; idx <= 3; idx++) {
                    if (descRaw === "PA_ETAT_ANALYSE_THEMATIQUETOP" + idx) {
                        shape.getText().setText(topThemParts[idx-1]);
                        appliquerMarkdownSurForme(shape);
                    }
                    if (descRaw === "PA_ETAT_ANALYSE_THEMATIQUEFLOP" + idx) {
                        shape.getText().setText(flopThemParts[idx-1]);
                        appliquerMarkdownSurForme(shape);
                    }
                    if (descRaw === "PA_ETAT_ANALYSE_MCTOP" + idx) {
                        shape.getText().setText(topSegParts[idx-1]);
                        appliquerMarkdownSurForme(shape);
                    }
                    if (descRaw === "PA_ETAT_ANALYSE_MCFLOP" + idx) {
                        shape.getText().setText(flopSegParts[idx-1]);
                        appliquerMarkdownSurForme(shape);
                    }
                }

                // Placeholders d'images concurrentes
                if (descRaw.indexOf("PLACEHOLDER_LOGO_") === 0) {
                    var imgUrl = null;
                    if (descRaw === "PLACEHOLDER_LOGO_CLIENT" && concurrenceData.client && concurrenceData.client.logoUrl) {
                        imgUrl = concurrenceData.client.logoUrl;
                    } else if (descRaw === "PLACEHOLDER_LOGO_LEADER" && concurrenceData.leader && concurrenceData.leader.logoUrl) {
                        imgUrl = concurrenceData.leader.logoUrl;
                    } else {
                        var m = descRaw.match(/PLACEHOLDER_LOGO_COMP(\d+)/);
                        if (m && m[1]) {
                            var idxComp = parseInt(m[1]) - 1;
                            if (concurrenceData.comps && concurrenceData.comps[idxComp] && concurrenceData.comps[idxComp].logoUrl) {
                                imgUrl = concurrenceData.comps[idxComp].logoUrl;
                            }
                        }
                    }

                    if (imgUrl) {
                        try {
                            var response = UrlFetchApp.fetch(imgUrl, { muteHttpExceptions: true });
                            if (response.getResponseCode() === 200) {
                                var blob = response.getBlob();
                                var newImg = slide.insertImage(blob, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                                newImg.setDescription(descRaw);
                            }
                            shape.remove();
                        } catch (errImg) {
                            Logger.log("Erreur chargement image " + imgUrl + " : " + errImg.message);
                            shape.remove();
                        }
                    } else {
                        shape.remove();
                    }
                }

                // --- Cas 2 : Remplacement local de placeholders dans le texte ({{...}}) ---
                if (shapeText && shapeText.indexOf("{{") !== -1) {
                    for (var key in replaceDict) {
                        if (shapeText.indexOf(key) !== -1) {
                            shape.getText().replaceText(key, String(replaceDict[key]));
                        }
                    }
                    appliquerMarkdownSurForme(shape); // S'il y a du markdown résiduel
                }

            });
            
            // Jauges (Cas 1 par description)
            var shapesForGauges = slide.getShapes();
            shapesForGauges.forEach(function(shape) {
                var descRaw = shape.getDescription() || "";
                var isTransacGauge = (descRaw === "JAUGE_TRANSAC_TOP10");
                var isInfoGauge = (descRaw === "JAUGE_INFO_TOP10");

                var targetGauge = isTransacGauge ? "transac" : (isInfoGauge ? "info" : null);

                if (targetGauge) {
                    var pct = (targetGauge === "transac") ? transacPctDec : infoPctDec;
                    var left = shape.getLeft();
                    var top = shape.getTop();
                    var width = shape.getWidth();

                    if (pct > 0) {
                        var fgShape = slide.insertShape(shape);
                        var fillWidth = Math.max(shape.getHeight(), width * pct);
                        fgShape.setWidth(fillWidth);
                        fgShape.setLeft(left);
                        fgShape.setTop(top);
                        fgShape.getFill().setSolidFill("#00b050");
                        fgShape.getBorder().setTransparent();
                        fgShape.getText().clear();
                        fgShape.setDescription("");
                    }
                    shape.getFill().setSolidFill("#f1f3f4");
                    shape.getBorder().setTransparent();
                    shape.getText().clear();
                    shape.setDescription("");
                }
            });
        });

        Logger.log("=== FIN : exporterPerformanceGlobalSlides ===");
        return { success: true, url: presentation.getUrl() };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT GLOBAL : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterFocusMotCleSlides() {
    try {
        Logger.log("=== DÉBUT : exporterFocusMotCleSlides ===");
        var props = getDatabaseData();
        var slideId = props['PA_CONF_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        var rawClientUrl = props['PA_FOCUS_MCCIBLE_URLCLIENT'] || "";
        var cleanClientUrl = rawClientUrl;
        if (rawClientUrl !== "" && rawClientUrl !== "-") {
            var matchPath = rawClientUrl.match(/^https?:\/\/[^\/]+(.*)$/i);
            cleanClientUrl = matchPath ? (matchPath[1] || "/") : rawClientUrl;
        }

        var rawCompUrl = props['PA_FOCUS_MCCIBLE_URLCONC'] || "";
        var cleanCompUrl = rawCompUrl;
        if (rawCompUrl !== "" && rawCompUrl !== "-") {
            cleanCompUrl = rawCompUrl.replace(/^https?:\/\//i, "");
        }
        
        var rawClientPos = props['PA_FOCUS_MCCIBLE_POSCLIENT'] || "";
        var formatedClientPos = (rawClientPos && rawClientPos !== "-") ? "Position " + rawClientPos : rawClientPos;
        var rawCompPos = props['PA_FOCUS_MCCIBLE_POSCONC'] || "";
        var formatedCompPos = (rawCompPos && rawCompPos !== "-") ? "Position " + rawCompPos : rawCompPos;

        var rawSv = props['PA_FOCUS_MCCIBLE_VOLUME'] || "";
        var formatedSv = rawSv ? rawSv + " rech./mois" : "";

        var simpleMapping = {
            'PA_FOCUS_SERP_ELEMENT_1': props['PA_FOCUS_SERP_ELEMENT_1'] || "",
            'PA_FOCUS_SERP_ELEMENT_2': props['PA_FOCUS_SERP_ELEMENT_2'] || "",
            'PA_FOCUS_SERP_ELEMENT_3': props['PA_FOCUS_SERP_ELEMENT_3'] || "",
            'PA_FOCUS_SERP_ELEMENT_4': props['PA_FOCUS_SERP_ELEMENT_4'] || "",
            'PA_FOCUS_INTENTION_TITRE': props['PA_FOCUS_INTENTION_TITRE'] || "",
            'PA_FOCUS_GAP_TITRE_1': props['PA_FOCUS_GAP_TITRE_1'] || "",
            'PA_FOCUS_GAP_TITRE_2': props['PA_FOCUS_GAP_TITRE_2'] || "",
            'PA_FOCUS_GAP_TITRE_3': props['PA_FOCUS_GAP_TITRE_3'] || "",
            'PA_FOCUS_MCCIBLE': props['PA_FOCUS_MCCIBLE'] || "",
            'PA_FOCUS_MCCIBLE_VOLUME': formatedSv,
            'PA_FOCUS_MCCIBLE_URLCLIENT': cleanClientUrl,
            'PA_FOCUS_MCCIBLE_POSCLIENT': formatedClientPos,
            'PA_FOCUS_MCCIBLE_URLCONC': cleanCompUrl,
            'PA_FOCUS_MCCIBLE_POSCONC': formatedCompPos
        };

        var altTextMapping = {
            'PA_FOCUS_SERP_ELEMENT_DESC_1': props['PA_FOCUS_SERP_ELEMENT_DESC_1'] || "",
            'PA_FOCUS_SERP_ELEMENT_DESC_2': props['PA_FOCUS_SERP_ELEMENT_DESC_2'] || "",
            'PA_FOCUS_SERP_ELEMENT_DESC_3': props['PA_FOCUS_SERP_ELEMENT_DESC_3'] || "",
            'PA_FOCUS_SERP_ELEMENT_DESC_4': props['PA_FOCUS_SERP_ELEMENT_DESC_4'] || "",
            'PA_FOCUS_INTENTION_DESC': props['PA_FOCUS_INTENTION_DESC'] || "",
            'PA_FOCUS_GAP_DESC_1': props['PA_FOCUS_GAP_DESC_1'] || "",
            'PA_FOCUS_GAP_DESC_2': props['PA_FOCUS_GAP_DESC_2'] || "",
            'PA_FOCUS_GAP_DESC_3': props['PA_FOCUS_GAP_DESC_3'] || "",
            'PA_FOCUS_RECO_1': props['PA_FOCUS_RECO_1'] || "",
            'PA_FOCUS_RECO_2': props['PA_FOCUS_RECO_2'] || "",
            'PA_FOCUS_RECO_3': props['PA_FOCUS_RECO_3'] || "",
            'PA_FOCUS_RECO_4': props['PA_FOCUS_RECO_4'] || ""
        };

        var placeholderMapping = {
            '{{focus_standard_texte_1}}': props['PA_FOCUS_STANDARD_TEXTE_1'] || "",
            '{{focus_standard_texte_2}}': props['PA_FOCUS_STANDARD_TEXTE_2'] || "",
            '{{focus_standard_texte_3}}': props['PA_FOCUS_STANDARD_TEXTE_3'] || "",
            '{{focus_semantique_texte_1}}': props['PA_FOCUS_SEMANTIQUE_TEXTE_1'] || "",
            '{{focus_semantique_texte_2}}': props['PA_FOCUS_SEMANTIQUE_TEXTE_2'] || "",
            '{{focus_semantique_texte_3}}': props['PA_FOCUS_SEMANTIQUE_TEXTE_3'] || ""
        };

        Logger.log("Parcours récursif des slides pour le Focus Mot-Clé (Groupes et Tableaux inclus)");

        slides.forEach(function(slide) {
            
            function processElement(element) {
                var type = element.getPageElementType();
                
                if (type === SlidesApp.PageElementType.GROUP) {
                    element.asGroup().getChildren().forEach(processElement);
                } else if (type === SlidesApp.PageElementType.TABLE) {
                    var table = element.asTable();
                    for (var r = 0; r < table.getNumRows(); r++) {
                        for (var c = 0; c < table.getNumColumns(); c++) {
                            processTextContainer(table.getCell(r, c), slide);
                        }
                    }
                } else if (type === SlidesApp.PageElementType.SHAPE) {
                    processTextContainer(element.asShape(), slide);
                } else if (type === SlidesApp.PageElementType.IMAGE) {
                    processTextContainer(element.asImage(), slide);
                }
            }

            function processTextContainer(element, currentSlide) {
                var isShape = (typeof element.getDescription === 'function');
                var descRaw = isShape ? (element.getDescription() || "") : "";
                var shapeText = "";

                try {
                    shapeText = element.getText().asString();
                } catch(e) {
                    return; 
                }

                if (isShape && descRaw !== "") {
                    if (simpleMapping[descRaw] !== undefined) {
                        element.getText().setText(String(simpleMapping[descRaw]));
                        return; 
                    }
                    
                    if (altTextMapping[descRaw] !== undefined) {
                        element.getText().setText(String(altTextMapping[descRaw]));
                        appliquerMarkdownSurForme(element);
                        return; 
                    }

                    var isSerpIcon = (descRaw === "PA_FOCUS_PLACEHOLDER_SERPELEMENT_1" || 
                                      descRaw === "PA_FOCUS_PLACEHOLDER_SERPELEMENT_2" || 
                                      descRaw === "PA_FOCUS_PLACEHOLDER_SERPELEMENT_3" || 
                                      descRaw === "PA_FOCUS_PLACEHOLDER_SERPELEMENT_4");

                    if (isSerpIcon && props[descRaw] && currentSlide) {
                        Logger.log("Détection du placeholder d'image : " + descRaw);
                        var featureName = props[descRaw].trim();
                        Logger.log("Valeur récupérée dans les propriétés : " + featureName);

                        if (featureName !== "") {
                            try {
                                var DRIVE_ICONS_MAPPING = {
                                    "organique": "19Fj-qair25NxZYx34lfqzbGkM5iAjVXY",
                                    "ads": "1pwUCDRXZ02ua0xuifRQaqnISmFnlAZsC",
                                    "featured": "1nh1r7ouYI6WktXkbIyBLRJbGADoxIaUO",
                                    "local": "1Wl2ZIe1REvW8_nWEAMbPFjYo2xE38VDA",
                                    "shopping": "13gt_1YTNJ_bJMTmPybOBc0x9N4_dy63p",
                                    "paa": "1frkC4wlrPqKwr6jxkcjEhE7HtWYf-6W-",
                                    "video": "1elbpXgnFxD4iSpSoHWFBswbUlYYnCkvL",
                                    "image": "1acgKroCoqPOy9rV2KnRwjdxk_fP_UIPh",
                                    "defaut": "18ILbiONR6N1gfikkFh-lMF1oTye45hje"
                                };

                                var finalFeature = DRIVE_ICONS_MAPPING[featureName] ? featureName : "defaut";
                                Logger.log("Feature finale utilisée : " + finalFeature);
                                
                                var fileId = DRIVE_ICONS_MAPPING[finalFeature];
                                Logger.log("ID Drive utilisé : " + fileId);
                                
                                var file = DriveApp.getFileById(fileId);
                                var pngBlob = file.getBlob();
                                Logger.log("Blob récupéré via DriveApp.");

                                var newImg = currentSlide.insertImage(pngBlob, element.getLeft(), element.getTop(), element.getWidth(), element.getHeight());
                                newImg.setDescription(descRaw);
                                element.remove();
                                Logger.log("Insertion réussie.");
                                return; 
                            } catch (errDrive) {
                                Logger.log("ERREUR CRITIQUE LORS DE L'INSERTION DE L'IMAGE DRIVE : " + errDrive.message);
                            }
                        }
                    }
                }

                if (shapeText) {
                    var hasReplaced = false;
                    for (var phKey in placeholderMapping) {
                        if (shapeText.indexOf(phKey) !== -1) {
                            element.getText().replaceAllText(phKey, placeholderMapping[phKey]);
                            hasReplaced = true;
                        }
                    }
                    if (hasReplaced) {
                        appliquerMarkdownSurForme(element);
                    }
                }
            }

            slide.getPageElements().forEach(processElement);
        });

        Logger.log("=== FIN : exporterFocusMotCleSlides ===");
        return { success: true, url: presentation.getUrl() };

    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT FOCUS : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterEtatLieuxTechniqueSlides() {
    try {
        Logger.log("=== DÉBUT : exporterEtatLieuxTechniqueSlides ===");
        var props = getDatabaseData();
        var slideId = props['PA_CONF_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        var ICON_IDS = {
            "BON": "1lwxjX4LJWDoNYb19qco0VK93EH1V_aaQ",
            "MOYEN": "1l-eMhlZ4eXu2zxzH-_D_ZdRbIWB3X7VB",
            "MAUVAIS": "1WCVH1kIsBu5oEG_nWP9fQsGS5JZ5aGgI",
            "INCONNU": "1bi8wj96QvF9EetPHPEkVztTEwZf5H8tS"
        };

        var textMapping = {
            'PA_TECH_TITRE': props['PA_TECH_TITRE'] || "",
            'PA_TECH_CRAWL_CONTENT_1': props['PA_TECH_CRAWL_CONTENT_1'] || "",
            'PA_TECH_CRAWL_CONTENT_2': props['PA_TECH_CRAWL_CONTENT_2'] || "",
            'PA_TECH_CRAWL_CONTENT_3': props['PA_TECH_CRAWL_CONTENT_3'] || "",
            'PA_TECH_INDEX_CONTENT_1': props['PA_TECH_INDEX_CONTENT_1'] || "",
            'PA_TECH_INDEX_CONTENT_2': props['PA_TECH_INDEX_CONTENT_2'] || "",
            'PA_TECH_INDEX_CONTENT_3': props['PA_TECH_INDEX_CONTENT_3'] || "",
            'PA_TECH_POS_CONTENT_1': props['PA_TECH_POS_CONTENT_1'] || "",
            'PA_TECH_POS_CONTENT_2': props['PA_TECH_POS_CONTENT_2'] || "",
            'PA_TECH_POS_CONTENT_3': props['PA_TECH_POS_CONTENT_3'] || ""
        };

        Logger.log("Parcours récursif des slides pour l'État des lieux technique...");

        slides.forEach(function(slide) {
            
            function processElement(element) {
                var type = element.getPageElementType();
                
                if (type === SlidesApp.PageElementType.GROUP) {
                    element.asGroup().getChildren().forEach(processElement);
                } else if (type === SlidesApp.PageElementType.TABLE) {
                    var table = element.asTable();
                    for (var r = 0; r < table.getNumRows(); r++) {
                        for (var c = 0; c < table.getNumColumns(); c++) {
                            processTextContainer(table.getCell(r, c), slide);
                        }
                    }
                } else if (type === SlidesApp.PageElementType.SHAPE) {
                    processTextContainer(element.asShape(), slide);
                } else if (type === SlidesApp.PageElementType.IMAGE) {
                    processTextContainer(element.asImage(), slide);
                }
            }

            function processTextContainer(element, currentSlide) {
                var isShape = (typeof element.getDescription === 'function');
                var descRaw = isShape ? (element.getDescription() || "") : "";

                if (isShape && descRaw !== "") {
                    // Remplacement des textes d'analyse
                    if (textMapping[descRaw] !== undefined) {
                        element.getText().setText(String(textMapping[descRaw]));
                        appliquerMarkdownSurForme(element);
                        return;
                    }

                    // Remplacement des icônes d'évaluation (BON, MOYEN, MAUVAIS, INCONNU)
                    var isIconPlaceholder = (
                        descRaw === "PA_TECH_CRAWL_CHECK_1" || descRaw === "PA_TECH_CRAWL_CHECK_2" || descRaw === "PA_TECH_CRAWL_CHECK_3" ||
                        descRaw === "PA_TECH_INDEX_CHECK_1" || descRaw === "PA_TECH_INDEX_CHECK_2" || descRaw === "PA_TECH_INDEX_CHECK_3" ||
                        descRaw === "PA_TECH_POS_CHECK_1" || descRaw === "PA_TECH_POS_CHECK_2" || descRaw === "PA_TECH_POS_CHECK_3"
                    );

                    if (isIconPlaceholder && currentSlide) {
                        var statusValue = props[descRaw] || "INCONNU";
                        Logger.log("Remplacement de l'icône " + descRaw + " par le statut : " + statusValue);
                        
                        var fileId = ICON_IDS[statusValue] || ICON_IDS["INCONNU"];
                        
                        try {
                            var file = DriveApp.getFileById(fileId);
                            var pngBlob = file.getBlob();
                            var newImg = currentSlide.insertImage(pngBlob, element.getLeft(), element.getTop(), element.getWidth(), element.getHeight());
                            newImg.setDescription(descRaw);
                            element.remove();
                        } catch (errDrive) {
                            Logger.log("ERREUR lors de l'insertion de l'icône Drive (" + statusValue + ") : " + errDrive.message);
                        }
                        return;
                    }
                }
            }

            slide.getPageElements().forEach(processElement);
        });

        Logger.log("=== FIN : exporterEtatLieuxTechniqueSlides ===");
        return { success: true, url: presentation.getUrl() };

    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT ETAT LIEUX TECHNIQUE : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterUXSlides() {
    try {
        Logger.log("=== DÉBUT : exporterUXSlides ===");
        var props = getDatabaseData();
        var slideId = props['PA_CONF_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        var ICON_IDS = {
            "BON": "1lwxjX4LJWDoNYb19qco0VK93EH1V_aaQ",
            "MOYEN": "1l-eMhlZ4eXu2zxzH-_D_ZdRbIWB3X7VB",
            "MAUVAIS": "1WCVH1kIsBu5oEG_nWP9fQsGS5JZ5aGgI",
            "INCONNU": "1bi8wj96QvF9EetPHPEkVztTEwZf5H8tS"
        };

        var elementsMapping = {};
        for (var i = 1; i <= 6; i++) {
            elementsMapping['PA_UX_ELEMENT_' + i] = props['PA_UX_ELEMENT_' + i] || "";
        }
        
        var recoMapping = {
            'PA_UX_TITRE': props['PA_UX_TITRE'] || "",
            'PA_UX_RECO_1': props['PA_UX_RECO_1'] || "",
            'PA_UX_RECO_2': props['PA_UX_RECO_2'] || ""
        };

        var clientViewportId = props['PA_UX_CLIENT_VIEWPORT'];
        var compViewportId = props['PA_UX_CONC_VIEWPORT'];

        Logger.log("Parcours récursif des slides pour l'UX...");

        slides.forEach(function(slide) {
            
            function processElement(element) {
                var type = element.getPageElementType();
                
                if (type === SlidesApp.PageElementType.GROUP) {
                    element.asGroup().getChildren().forEach(processElement);
                } else if (type === SlidesApp.PageElementType.TABLE) {
                    var table = element.asTable();
                    for (var r = 0; r < table.getNumRows(); r++) {
                        for (var c = 0; c < table.getNumColumns(); c++) {
                            processTextContainer(table.getCell(r, c), slide);
                        }
                    }
                } else if (type === SlidesApp.PageElementType.SHAPE) {
                    processTextContainer(element.asShape(), slide);
                } else if (type === SlidesApp.PageElementType.IMAGE) {
                    processTextContainer(element.asImage(), slide);
                }
            }

            function processTextContainer(element, currentSlide) {
                var isShape = (typeof element.getDescription === 'function');
                var descRaw = isShape ? (element.getDescription() || "") : "";

                if (isShape && descRaw !== "") {
                    // Éléments UX
                    if (elementsMapping[descRaw] !== undefined) {
                        element.getText().setText(String(elementsMapping[descRaw]));
                        return;
                    }

                    // Recommandations UX
                    if (recoMapping[descRaw] !== undefined) {
                        element.getText().setText(String(recoMapping[descRaw]));
                        appliquerMarkdownSurForme(element);
                        return;
                    }

                    // Icônes d'évaluation
                    var isIconPlaceholder = descRaw.indexOf("PA_UX_CLIENT_CHECK_") === 0 || descRaw.indexOf("PA_UX_CONC_CHECK_") === 0;
                    if (isIconPlaceholder && currentSlide) {
                        var statusValue = props[descRaw] || "INCONNU";
                        Logger.log("Remplacement de l'icône " + descRaw + " par le statut : " + statusValue);
                        
                        var fileId = ICON_IDS[statusValue] || ICON_IDS["INCONNU"];
                        
                        try {
                            var file = DriveApp.getFileById(fileId);
                            var pngBlob = file.getBlob();
                            var newImg = currentSlide.insertImage(pngBlob, element.getLeft(), element.getTop(), element.getWidth(), element.getHeight());
                            newImg.setDescription(descRaw);
                            element.remove();
                        } catch (errDrive) {
                            Logger.log("ERREUR lors de l'insertion de l'icône Drive (" + statusValue + ") : " + errDrive.message);
                        }
                        return;
                    }

                    // Captures d'écran
                    if (descRaw === "PA_UX_PLACEHOLDER_CLIENT" && clientViewportId && currentSlide) {
                        Logger.log("Insertion capture client : " + clientViewportId);
                        try {
                            var fileClient = DriveApp.getFileById(clientViewportId);
                            var imgClient = currentSlide.insertImage(fileClient.getBlob(), element.getLeft(), element.getTop(), element.getWidth(), element.getHeight());
                            imgClient.setDescription(descRaw);
                            element.remove();
                        } catch (e) {
                            Logger.log("Erreur insertion capture client : " + e.message);
                        }
                        return;
                    }

                    if (descRaw === "PA_UX_PLACEHOLDER_CONC" && compViewportId && currentSlide) {
                        Logger.log("Insertion capture concurrent : " + compViewportId);
                        try {
                            var fileComp = DriveApp.getFileById(compViewportId);
                            var imgComp = currentSlide.insertImage(fileComp.getBlob(), element.getLeft(), element.getTop(), element.getWidth(), element.getHeight());
                            imgComp.setDescription(descRaw);
                            element.remove();
                        } catch (e) {
                            Logger.log("Erreur insertion capture concurrent : " + e.message);
                        }
                        return;
                    }
                }
            }

            slide.getPageElements().forEach(processElement);
        });

        Logger.log("=== FIN : exporterUXSlides ===");
        return { success: true, url: presentation.getUrl() };

    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT UX : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterEditorialSlides(concurrenceDataEdito, titreSlide, donneesBlog) {
    try {
        Logger.log("=== DÉBUT : exporterEditorialSlides ===");
        var props = getDatabaseData();
        var slideId = props['PA_CONF_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        function safeNum(val) {
            return (val !== null && val !== undefined && !isNaN(val)) ? Math.round(val).toLocaleString('fr-FR') : "-";
        }

        function colorerBlogEdito(shape, valeur) {
            try {
                var color = (valeur && valeur.toString().toLowerCase() === "oui") ? "#02b050" : "#ff0000";
                shape.getFill().setSolidFill(color);
                shape.getText().getTextStyle().setForegroundColor("#ffffff").setBold(true);
            } catch (e) {
                Logger.log("Erreur couleur blog : " + e.message);
            }
        }

        var mappingComp = {};
        
        // Titres
        mappingComp['PA_EDITO_TITRE_CONC'] = titreSlide || props['PA_EDITO_TITRE_CONC'] || "";
        mappingComp['PA_EDITO_TITRE_THEMATIQUE'] = props['PA_EDITO_TITRE_THEMATIQUE'] || "";
        
        if (concurrenceDataEdito) {
            // Paysage concurrentiel
            if (concurrenceDataEdito.client) {
                mappingComp['PA_EDITO_NOM_CLIENT'] = concurrenceDataEdito.client.name;
                mappingComp['PA_EDITO_TOP10_CLIENT'] = safeNum(concurrenceDataEdito.client.top10);
                mappingComp['PA_EDITO_PAGES_CLIENT'] = safeNum(concurrenceDataEdito.client.pages);
                mappingComp['PA_EDITO_BLOG_CLIENT'] = donneesBlog.client || "Non";
            }
            if (concurrenceDataEdito.leader) {
                mappingComp['PA_EDITO_NOM_LEADER'] = concurrenceDataEdito.leader.name;
                mappingComp['PA_EDITO_TOP10_LEADER'] = safeNum(concurrenceDataEdito.leader.top10);
                mappingComp['PA_EDITO_PAGES_LEADER'] = safeNum(concurrenceDataEdito.leader.pages);
                mappingComp['PA_EDITO_BLOG_LEADER'] = donneesBlog.leader || "Non";
            }
            for (var c = 1; c <= 4; c++) {
                var comp = concurrenceDataEdito.comps && concurrenceDataEdito.comps[c-1] ? concurrenceDataEdito.comps[c-1] : null;
                if (comp) {
                    mappingComp['PA_EDITO_NOM_COMP' + c] = comp.name;
                    mappingComp['PA_EDITO_TOP10_COMP' + c] = safeNum(comp.top10);
                    mappingComp['PA_EDITO_PAGES_COMP' + c] = safeNum(comp.pages);
                    mappingComp['PA_EDITO_BLOG_CONC' + c] = donneesBlog['comp' + c] || "Non";
                }
            }
            
            // Pistes Éditoriales (Thématiques)
            var selectionJson = PropertiesService.getScriptProperties().getProperty('PA_DIAGNOSTIC_SELECTION');
            var selection = selectionJson ? JSON.parse(selectionJson) : [];
            var pistesEdito = [];
            try {
                var diag = genererDiagnostic(selection);
                if (diag && diag.pistesEdito) pistesEdito = diag.pistesEdito;
            } catch (e) { Logger.log("Erreur diagnostic : " + e.message); }
            
            for (var i = 1; i <= 3; i++) {
                var piste = pistesEdito[i - 1];
                if (piste) {
                    // Thématique avec retour à la ligne
                    mappingComp['PA_EDITO_THEMATIQUE_' + i] = piste.tsKey ? piste.tsKey.replace(" > ", "\n> ") : "";
                    mappingComp['PA_EDITO_NOM_CONC_CONTENU_' + i] = piste.entite || "";
                    mappingComp['PA_EDITO_URL_CONTENU_' + i] = piste.url || "";
                    mappingComp['PA_EDITO_NOM_CONTENU_' + i] = props['PA_EDITO_NOM_CONTENU_' + i] || "";
                    
                    // Récupération directe du bloc Top 10 formaté par le front-end
                    var dataTop10 = props['PA_EDITO_DATA_TOP10_' + i];
                    mappingComp['PA_EDITO_DATA_TOP10_' + i] = dataTop10 ? dataTop10.replace(/&#10;/g, "\n") : "-";
                } else {
                    mappingComp['PA_EDITO_THEMATIQUE_' + i] = "-";
                    mappingComp['PA_EDITO_NOM_CONC_CONTENU_' + i] = "-";
                    mappingComp['PA_EDITO_URL_CONTENU_' + i] = "-";
                    mappingComp['PA_EDITO_NOM_CONTENU_' + i] = "-";
                    mappingComp['PA_EDITO_DATA_TOP10_' + i] = "-";
                }
            }
        }

        var propsToSave = {};
        for (var k in mappingComp) { propsToSave[k] = String(mappingComp[k]); }

        propsToSave["PLACEHOLDER_LOGO_CLIENT_EDITO"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_LEADER_EDITO"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP1_EDITO"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP2_EDITO"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP3_EDITO"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP4_EDITO"] = "IMAGE";
        
        propsToSave["PLACEHOLDER_LOGO_COMP1_EDITO_CONTENU"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP2_EDITO_CONTENU"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP3_EDITO_CONTENU"] = "IMAGE";

        setDatabaseData(propsToSave);

        Logger.log("Parcours récursif des slides pour Performance Éditoriale (Cas 1 Uniquement, avec Groupes)");
        slides.forEach(function(slide) {
            
            function processElement(element) {
                var type = element.getPageElementType();
                
                if (type === SlidesApp.PageElementType.GROUP) {
                    element.asGroup().getChildren().forEach(processElement);
                } else if (type === SlidesApp.PageElementType.TABLE) {
                    var table = element.asTable();
                    for (var r = 0; r < table.getNumRows(); r++) {
                        for (var c = 0; c < table.getNumColumns(); c++) {
                            processTextContainer(table.getCell(r, c), slide);
                        }
                    }
                } else if (type === SlidesApp.PageElementType.SHAPE) {
                    processTextContainer(element.asShape(), slide);
                } else if (type === SlidesApp.PageElementType.IMAGE) {
                    processTextContainer(element.asImage(), slide);
                }
            }

            function processTextContainer(element, currentSlide) {
                var isShape = (typeof element.getDescription === 'function');
                var descRaw = isShape ? (element.getDescription() || "") : "";

                if (isShape && descRaw !== "") {
                    
                    // 1. Remplacement des textes (Titres, Thématiques, Données, Tags Blog)
                    if (mappingComp[descRaw] !== undefined) {
                        element.getText().setText(mappingComp[descRaw].toString());

                        if (descRaw.indexOf("PA_EDITO_BLOG_") === 0) {
                            colorerBlogEdito(element, mappingComp[descRaw]);
                        }
                        
                        // Application systématique de la transformation Markdown sur le texte final
                        appliquerMarkdownSurForme(element);
                        
                        return; // On stoppe ici pour cette forme
                    }

                    // 2. Remplacement des images (Logos/Favicons)
                    if (descRaw.indexOf("PLACEHOLDER_LOGO_") === 0 && descRaw.indexOf("_EDITO") !== -1) {
                        var imgUrl = null;
                        if (descRaw.indexOf("_EDITO_CONTENU") !== -1) {
                            var mc = descRaw.match(/PLACEHOLDER_LOGO_COMP(\d+)_EDITO_CONTENU/);
                            if (mc && mc[1]) {
                                var domainUrl = mappingComp['PA_EDITO_URL_CONTENU_' + parseInt(mc[1])];
                                if (domainUrl && domainUrl !== "-") {
                                    var cleanDomain = domainUrl.replace(/^(?:https?:\/\/)?(?:www\.)?/i, "").split('/')[0];
                                    imgUrl = 'https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&url=https://' + cleanDomain + '&size=128';
                                }
                            }
                        } else {
                            if (descRaw === "PLACEHOLDER_LOGO_CLIENT_EDITO" && concurrenceDataEdito.client && concurrenceDataEdito.client.logoUrl) imgUrl = concurrenceDataEdito.client.logoUrl;
                            else if (descRaw === "PLACEHOLDER_LOGO_LEADER_EDITO" && concurrenceDataEdito.leader && concurrenceDataEdito.leader.logoUrl) imgUrl = concurrenceDataEdito.leader.logoUrl;
                            else {
                                var m = descRaw.match(/PLACEHOLDER_LOGO_COMP(\d+)_EDITO/);
                                if (m && m[1] && concurrenceDataEdito.comps && concurrenceDataEdito.comps[parseInt(m[1]) - 1]) imgUrl = concurrenceDataEdito.comps[parseInt(m[1]) - 1].logoUrl;
                            }
                        }

                        if (imgUrl) {
                            try {
                                var response = UrlFetchApp.fetch(imgUrl, { muteHttpExceptions: true });
                                if (response.getResponseCode() === 200) {
                                    currentSlide.insertImage(response.getBlob(), element.getLeft(), element.getTop(), element.getWidth(), element.getHeight()).setDescription(descRaw);
                                }
                                element.remove();
                            } catch (e) { element.remove(); }
                        } else { element.remove(); }
                        return; // On stoppe ici pour cette forme
                    }
                }
            }

            slide.getPageElements().forEach(processElement);
        });

        Logger.log("=== FIN : exporterEditorialSlides ===");
        return { success: true, url: presentation.getUrl() };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT EDITORIAL : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterContactsSlides() {
    try {
        Logger.log("=== DÉBUT : exporterContactsSlides ===");
        var props = getDatabaseData();
        var slideId = props['PA_CONF_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        var textReplaceMapping = {};
        
        // Commerciaux
        textReplaceMapping['{{NOM_COM}}'] = props['nom_com'] || "";
        textReplaceMapping['{{POSTE_COM}}'] = props['poste_com'] || "";
        textReplaceMapping['{{EMAIL_COM}}'] = props['email_com'] || "";
        
        // Consultant 1
        textReplaceMapping['{{NOM_CONS1}}'] = props['nom_cons1'] || "";
        textReplaceMapping['{{POSTE_CONS1}}'] = props['poste_cons1'] || "";
        textReplaceMapping['{{EMAIL_CONS1}}'] = props['email_cons1'] || "";
        
        // Consultant 2 (Optionnel)
        var nomCons2 = props['nom_cons2'] || "";
        var imgCons2 = props['PLACEHOLDER_CONTACT_CONS2'] || "";
        var isCons2Valid = (nomCons2 !== "" && imgCons2 !== "");

        if (isCons2Valid) {
            textReplaceMapping['{{NOM_CONS2}}'] = nomCons2;
            textReplaceMapping['{{POSTE_CONS2}}'] = props['poste_cons2'] || "";
            textReplaceMapping['{{EMAIL_CONS2}}'] = props['email_cons2'] || "";
        } else {
            textReplaceMapping['{{NOM_CONS2}}'] = "";
            textReplaceMapping['{{POSTE_CONS2}}'] = "";
            textReplaceMapping['{{EMAIL_CONS2}}'] = "";
        }

        slides.forEach(function(slide) {
            var shapesText = slide.getShapes();
            shapesText.forEach(function(shape) {
                var shapeText = "";
                try {
                    shapeText = shape.getText().asString();
                } catch(e) {}

                // Sécurité : suppression de la forme texte si cons2 n'est pas valide
                if (!isCons2Valid && shapeText && (shapeText.indexOf("{{NOM_CONS2}}") !== -1 || shapeText.indexOf("{{POSTE_CONS2}}") !== -1 || shapeText.indexOf("{{EMAIL_CONS2}}") !== -1)) {
                    shape.remove();
                    return;
                }

                if (shapeText && shapeText.indexOf("{{") !== -1) {
                    for (var key in textReplaceMapping) {
                        if (shapeText.indexOf(key) !== -1) {
                            var value = textReplaceMapping[key];
                            var textRange = shape.getText();
                            
                            // Remplacement global du tag par la valeur
                            textRange.replaceAllText(key, value);
                            
                            // Si c'est un email, on repère ses nouvelles coordonnées textuelles pour appliquer le lien
                            if (key.indexOf('EMAIL') !== -1 && value !== "") {
                                var newTextStr = textRange.asString();
                                var searchIndex = 0;
                                while ((searchIndex = newTextStr.indexOf(value, searchIndex)) !== -1) {
                                    var emailRange = textRange.getRange(searchIndex, searchIndex + value.length);
                                    emailRange.getTextStyle().setLinkUrl("mailto:" + value);
                                    searchIndex += value.length;
                                }
                            }
                            
                            // Actualisation de shapeText pour l'itération des variables suivantes
                            shapeText = textRange.asString();
                        }
                    }
                }
            });

            var shapesImg = slide.getShapes();
            shapesImg.forEach(function(shape) {
                var descRaw = shape.getDescription() || "";

                var targetDriveId = null;
                var shouldRemove = false;

                if (descRaw === "PLACEHOLDER_CONTACT_COM") {
                    targetDriveId = props['PLACEHOLDER_CONTACT_COM'];
                } else if (descRaw === "PLACEHOLDER_CONTACT_CONS1") {
                    targetDriveId = props['PLACEHOLDER_CONTACT_CONS1'];
                } else if (descRaw === "PLACEHOLDER_CONTACT_CONS2") {
                    if (isCons2Valid) {
                        targetDriveId = imgCons2;
                    } else {
                        shouldRemove = true;
                    }
                }

                if (shouldRemove) {
                    shape.remove();
                } else if (targetDriveId && targetDriveId !== "ID_DRIVE_A_REMPLIR") {
                    try {
                        var file = DriveApp.getFileById(targetDriveId);
                        var blob = file.getBlob();
                        shape.replaceWithImage(blob, true);
                    } catch (e) {
                        Logger.log("Erreur remplacement image contact pour " + descRaw + " : " + e.message);
                    }
                }
            });
        });

        Logger.log("=== FIN : exporterContactsSlides ===");
        return { success: true, url: presentation.getUrl() };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT CONTACTS : " + e.message);
        return { success: false, error: e.message };
    }
}
