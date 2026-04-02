function exporterSlideBesoinSolution(texteBesoin, texteSolution) {
    try {
        Logger.log("=== DÉBUT : exporterSlideBesoinSolution ===");
        var props = getDatabaseData();
        var slideId = props['PA_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();
        
        var tagsTrouves = 0;
        
        // Formatage : on ajoute un saut de ligne vide entre chaque puce pour aérer dans Slides
        var slideTexteBesoin = texteBesoin.replace(/\n/g, '\n\n');
        var slideTexteSolution = texteSolution.replace(/\n/g, '\n\n');
        
        Logger.log("Recherche des tags TAG_SLIDE_BESOIN et TAG_SLIDE_SOLUTION via description");

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();
            
            shapes.forEach(function(shape) {
                var descRaw = shape.getDescription() || "";
                var targetKey = null;

                // Détection via le texte alternatif en majuscule UNIQUEMENT (Cas 1)
                if (descRaw === "TAG_SLIDE_BESOIN") {
                    targetKey = "besoin";
                } else if (descRaw === "TAG_SLIDE_SOLUTION") {
                    targetKey = "solution";
                }

                if (targetKey === "besoin") {
                    Logger.log("Forme cible 'besoin' trouvée, écrasement du texte");
                    shape.getText().setText(slideTexteBesoin);
                    tagsTrouves++;
                } else if (targetKey === "solution") {
                    Logger.log("Forme cible 'solution' trouvée, écrasement du texte");
                    shape.getText().setText(slideTexteSolution);
                    tagsTrouves++;
                }
            });
        });
        
        Logger.log("Tags trouvés et remplacés : " + tagsTrouves);
        Logger.log("=== FIN : exporterSlideBesoinSolution ===");
        
        if (tagsTrouves === 0) {
            return { success: false, error: "Les tags 'TAG_SLIDE_BESOIN' et 'TAG_SLIDE_SOLUTION' n'ont pas été trouvés dans le texte alternatif de la présentation." };
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
        var slideId = props['PA_SLIDE_ID'];
        
        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        // Fonction utilitaire non destructive pour *mot* (gras orange)
        function formatRichTextSemrush(shape) {
            var textRange = shape.getText();
            var textStr = textRange.asString();
            var regex = /\*([^*]+)\*/g;
            var matches = [];
            var match;
            
            while ((match = regex.exec(textStr)) !== null) {
                matches.push({
                    start: match.index,
                    text: match[1],
                    length: match[0].length
                });
            }
            
            for (var i = matches.length - 1; i >= 0; i--) {
                var m = matches[i];
                var endAsteriskIndex = m.start + m.text.length + 1;
                
                // Effacer l'astérisque final
                textRange.getRange(endAsteriskIndex, endAsteriskIndex + 1).clear();
                // Effacer l'astérisque initial
                textRange.getRange(m.start, m.start + 1).clear();
                
                // Appliquer le style sur le texte restant
                var styledRange = textRange.getRange(m.start, m.start + m.text.length);
                styledRange.getTextStyle().setBold(true).setForegroundColor("#f67604");
            }
        }

        Logger.log("Parcours des slides pour l'analyse Semrush en cours");

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();
            
            shapes.forEach(function(shape) {
                var descRaw = shape.getDescription() || "";

                // Cas 1 : Description correspondante -> Écrasement total
                if (descRaw === "TITRE_SLIDE_SEMRUSH") {
                    Logger.log("Remplacement du titre TITRE_SLIDE_SEMRUSH");
                    shape.getText().setText(titre);
                }
                
                if (descRaw === "ANALYSE_SEMRUSH_MOT_CLE") {
                    Logger.log("Remplacement et formatage ANALYSE_SEMRUSH_MOT_CLE");
                    shape.getText().setText(texteKw);
                    formatRichTextSemrush(shape);
                }
                
                if (descRaw === "ANALYSE_SEMRUSH_TRAFIC") {
                    Logger.log("Remplacement et formatage ANALYSE_SEMRUSH_TRAFIC");
                    shape.getText().setText(texteTrafic);
                    formatRichTextSemrush(shape);
                }

                // Cas 1 sur Image (Placeholders)
                if (descRaw === "PLACEHOLDER_ANALYSE_SEMRUSH_MOT_CLE") {
                    Logger.log("Remplacement image mots-clés");
                    var blobKw = Utilities.newBlob(Utilities.base64Decode(imgKwB64), imgKwMime, "kw.png");
                    var newImageKw = slide.insertImage(blobKw, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                    newImageKw.setDescription(descRaw);
                    shape.remove();
                }
                
                if (descRaw === "PLACEHOLDER_ANALYSE_SEMRUSH_TRAFIC") {
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
        var slideId = props['PA_SLIDE_ID'];

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
            mappingComp['TITRE_SLIDE_CONCURRENCE'] = "L'environnement concurrentiel de " + (concurrenceData.client ? concurrenceData.client.name : "");
            if (concurrenceData.client) {
                mappingComp['NOM_CLIENT'] = concurrenceData.client.name;
                mappingComp['VALEUR_TOP10_CLIENT'] = safeNum(concurrenceData.client.top10);
                mappingComp['VALEUR_PAGES_CLIENT'] = safeNum(concurrenceData.client.pages);
            }
            if (concurrenceData.leader) {
                mappingComp['NOM_LEADER'] = concurrenceData.leader.name;
                mappingComp['VALEUR_TOP10_LEADER'] = safeNum(concurrenceData.leader.top10);
                mappingComp['VALEUR_PAGES_LEADER'] = safeNum(concurrenceData.leader.pages);
            }
            for (var c = 1; c <= 4; c++) {
                var comp = concurrenceData.comps && concurrenceData.comps[c-1] ? concurrenceData.comps[c-1] : null;
                if (comp) {
                    mappingComp['NOM_COMP' + c] = comp.name;
                    mappingComp['VALEUR_TOP10_COMP' + c] = safeNum(comp.top10);
                    mappingComp['VALEUR_PAGES_COMP' + c] = safeNum(comp.pages);
                }
            }
        }

        var mapping = {
            'MOTCLE_CLIENT_GLOBAL': (clientKpi.posAll || 0).toLocaleString('fr-FR'),
            'MOTCLE_CLIENT_TOP3': (clientKpi.top3 || 0).toLocaleString('fr-FR'),
            'MOTCLE_CLIENT_TOP10': (clientKpi.top10 || 0).toLocaleString('fr-FR'),
            'MOTCLE_CLIENT_URL': (clientKpi.urlsCount || 0).toLocaleString('fr-FR'),
            'MOTCLE_CLIENT_TRANSAC': (intentStats.transac.top100 || 0).toLocaleString('fr-FR'),
            'MOTCLE_CLIENT_INFO': (intentStats.info.top100 || 0).toLocaleString('fr-FR'),
            'MOTCLE_CLIENT_TRANSAC_TOP10': (intentStats.transac.top10 || 0).toLocaleString('fr-FR'),
            'MOTCLE_CLIENT_INFO_TOP10': (intentStats.info.top10 || 0).toLocaleString('fr-FR'),
            'MOTCLE_CLIENT_TRANSAC_PCT': Math.round(transacPctDec * 100) + "%",
            'MOTCLE_CLIENT_INFO_PCT': Math.round(infoPctDec * 100) + "%"
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
            propsToSave["ANALYSE_THEMATIQUETOP_CLIENT_" + idx] = topThemParts[idx-1];
            propsToSave["ANALYSE_THEMATIQUEFLOP_CLIENT_" + idx] = flopThemParts[idx-1];
            propsToSave["ANALYSE_MCTOP_CLIENT_" + idx] = topSegParts[idx-1];
            propsToSave["ANALYSE_MCFLOP_CLIENT_" + idx] = flopSegParts[idx-1];
        }

        propsToSave["PLACEHOLDER_LOGO_CLIENT"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_LEADER"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP1"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP2"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP3"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP4"] = "IMAGE";

        setDatabaseData(propsToSave);

        // Fonction utilitaire non destructive pour *mot* (gras orange)
        function formatRichTextGlobal(shape) {
            var textRange = shape.getText();
            var textStr = textRange.asString();
            var regex = /\*([^*]+)\*/g;
            var matches = [];
            var match;
            
            while ((match = regex.exec(textStr)) !== null) {
                matches.push({
                    start: match.index,
                    text: match[1],
                    length: match[0].length
                });
            }
            
            for (var i = matches.length - 1; i >= 0; i--) {
                var m = matches[i];
                var endAsteriskIndex = m.start + m.text.length + 1;
                
                textRange.getRange(endAsteriskIndex, endAsteriskIndex + 1).clear();
                textRange.getRange(m.start, m.start + 1).clear();
                
                var styledRange = textRange.getRange(m.start, m.start + m.text.length);
                styledRange.getTextStyle().setBold(true).setForegroundColor("#f67604");
            }
        }

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
                if (descRaw === "TITRE_SLIDE_THEMATIQUETOP_CLIENT" && iaData && iaData.titreTopThematiques) {
                    shape.getText().setText(iaData.titreTopThematiques);
                }
                if (descRaw === "TITRE_SLIDE_THEMATIQUEFLOP_CLIENT" && iaData && iaData.titreFlopThematiques) {
                    shape.getText().setText(iaData.titreFlopThematiques);
                }
                if (descRaw === "TITRE_SLIDE_MCTOP_CLIENT" && iaData && iaData.titreTopSegments) {
                    shape.getText().setText(iaData.titreTopSegments);
                }
                if (descRaw === "TITRE_SLIDE_MCFLOP_CLIENT" && iaData && iaData.titreFlopSegments) {
                    shape.getText().setText(iaData.titreFlopSegments);
                }

                // Analyses IA découpées
                for (var idx = 1; idx <= 3; idx++) {
                    if (descRaw === "ANALYSE_THEMATIQUETOP_CLIENT_" + idx) {
                        shape.getText().setText(topThemParts[idx-1]);
                        formatRichTextGlobal(shape);
                    }
                    if (descRaw === "ANALYSE_THEMATIQUEFLOP_CLIENT_" + idx) {
                        shape.getText().setText(flopThemParts[idx-1]);
                        formatRichTextGlobal(shape);
                    }
                    if (descRaw === "ANALYSE_MCTOP_CLIENT_" + idx) {
                        shape.getText().setText(topSegParts[idx-1]);
                        formatRichTextGlobal(shape);
                    }
                    if (descRaw === "ANALYSE_MCFLOP_CLIENT_" + idx) {
                        shape.getText().setText(flopSegParts[idx-1]);
                        formatRichTextGlobal(shape);
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
                    formatRichTextGlobal(shape); // S'il y a du markdown résiduel
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
        var slideId = props['PA_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        var rawClientUrl = props['TARGET_URL_CLIENT'] || "";
        var cleanClientUrl = rawClientUrl;
        if (rawClientUrl !== "" && rawClientUrl !== "-") {
            var matchPath = rawClientUrl.match(/^https?:\/\/[^\/]+(.*)$/i);
            cleanClientUrl = matchPath ? (matchPath[1] || "/") : rawClientUrl;
        }

        var rawCompUrl = props['TARGET_URL_CONCURRENT'] || "";
        var cleanCompUrl = rawCompUrl;
        if (rawCompUrl !== "" && rawCompUrl !== "-") {
            cleanCompUrl = rawCompUrl.replace(/^https?:\/\//i, "");
        }
        
        var rawClientPos = props['TARGET_KW_CLIENT_POS'] || "";
        var formatedClientPos = (rawClientPos && rawClientPos !== "-") ? "Position " + rawClientPos : rawClientPos;
        var rawCompPos = props['TARGET_KW_CONCURRENT_POS'] || "";
        var formatedCompPos = (rawCompPos && rawCompPos !== "-") ? "Position " + rawCompPos : rawCompPos;

        var rawSv = props['TARGET_KW_SV'] || "";
        var formatedSv = rawSv ? rawSv + " rech./mois" : "";

        var simpleMapping = {
            'SERP_ELEMENT_TITRE_1': props['SERP_ELEMENT_TITRE_1'] || "",
            'SERP_ELEMENT_TITRE_2': props['SERP_ELEMENT_TITRE_2'] || "",
            'SERP_ELEMENT_TITRE_3': props['SERP_ELEMENT_TITRE_3'] || "",
            'SERP_ELEMENT_TITRE_4': props['SERP_ELEMENT_TITRE_4'] || "",
            'FOCUS_INTENTION_TITRE': props['FOCUS_INTENTION_TITRE'] || "",
            'FOCUS_GAP_TITRE_1': props['FOCUS_GAP_TITRE_1'] || "",
            'FOCUS_GAP_TITRE_2': props['FOCUS_GAP_TITRE_2'] || "",
            'FOCUS_GAP_TITRE_3': props['FOCUS_GAP_TITRE_3'] || "",
            'TARGET_KW': props['TARGET_KW'] || "",
            'TARGET_KW_SV': formatedSv,
            'TARGET_URL_CLIENT': cleanClientUrl,
            'TARGET_KW_CLIENT_POS': formatedClientPos,
            'TARGET_URL_CONCURRENT': cleanCompUrl,
            'TARGET_KW_CONCURRENT_POS': formatedCompPos
        };

        var altTextMapping = {
            'SERP_ELEMENT_DESC_1': props['SERP_ELEMENT_DESC_1'] || "",
            'SERP_ELEMENT_DESC_2': props['SERP_ELEMENT_DESC_2'] || "",
            'SERP_ELEMENT_DESC_3': props['SERP_ELEMENT_DESC_3'] || "",
            'SERP_ELEMENT_DESC_4': props['SERP_ELEMENT_DESC_4'] || "",
            'FOCUS_INTENTION_DESC': props['FOCUS_INTENTION_DESC'] || "",
            'FOCUS_GAP_DESC_1': props['FOCUS_GAP_DESC_1'] || "",
            'FOCUS_GAP_DESC_2': props['FOCUS_GAP_DESC_2'] || "",
            'FOCUS_GAP_DESC_3': props['FOCUS_GAP_DESC_3'] || "",
            'FOCUS_RECO_1': props['FOCUS_RECO_1'] || "",
            'FOCUS_RECO_2': props['FOCUS_RECO_2'] || "",
            'FOCUS_RECO_3': props['FOCUS_RECO_3'] || "",
            'FOCUS_RECO_4': props['FOCUS_RECO_4'] || ""
        };

        var placeholderMapping = {
            '{{focus_standard_texte_1}}': props['focus_standard_texte_1'] || "",
            '{{focus_standard_texte_2}}': props['focus_standard_texte_2'] || "",
            '{{focus_standard_texte_3}}': props['focus_standard_texte_3'] || "",
            '{{focus_semantique_texte_1}}': props['focus_semantique_texte_1'] || "",
            '{{focus_semantique_texte_2}}': props['focus_semantique_texte_2'] || "",
            '{{focus_semantique_texte_3}}': props['focus_semantique_texte_3'] || ""
        };

        function formatRichTextFocus(element) {
            try {
                var textRange = element.getText();
                var textStr = textRange.asString();
                var regex = /\*\*([^*]+)\*\*/g;
                var matches = [];
                var match;

                while ((match = regex.exec(textStr)) !== null) {
                    matches.push({
                        start: match.index,
                        text: match[1],
                        length: match[0].length
                    });
                }
                
                for (var i = matches.length - 1; i >= 0; i--) {
                    var m = matches[i];
                    var endDoubleAst = m.start + m.text.length + 2;
                    
                    textRange.getRange(endDoubleAst, endDoubleAst + 2).clear();
                    textRange.getRange(m.start, m.start + 2).clear();

                    var styledRange = textRange.getRange(m.start, m.start + m.text.length);
                    styledRange.getTextStyle().setBold(true);
                }
            } catch(e) {}
        }

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
                        formatRichTextFocus(element);
                        return; 
                    }

                    var isSerpIcon = (descRaw === "PLACEHOLDER_SERPELEMENT_1" || 
                                      descRaw === "PLACEHOLDER_SERPELEMENT_2" || 
                                      descRaw === "PLACEHOLDER_SERPELEMENT_3" || 
                                      descRaw === "PLACEHOLDER_SERPELEMENT_4");

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
                        formatRichTextFocus(element);
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
        var slideId = props['PA_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        var ICON_IDS = {
            "BON": "1lwxjX4LJWDoNYb19qco0VK93EH1V_aaQ",
            "MOYEN": "1l-eMhlZ4eXu2zxzH-_D_ZdRbIWB3X7VB",
            "MAUVAIS": "1WCVH1kIsBu5oEG_nWP9fQsGS5JZ5aGgI",
            "INCONNU": "1bi8wj96QvF9EetPHPEkVztTEwZf5H8tS"
        };

        function formatRichTextTech(shape) {
            try {
                var textRange = shape.getText();
                var textStr = textRange.asString();
                var regex = /\*\*([^*]+)\*\*/g;
                var matches = [];
                var match;

                while ((match = regex.exec(textStr)) !== null) {
                    matches.push({
                        start: match.index,
                        text: match[1],
                        length: match[0].length
                    });
                }
                
                for (var i = matches.length - 1; i >= 0; i--) {
                    var m = matches[i];
                    var endDoubleAst = m.start + m.text.length + 2;
                    
                    textRange.getRange(endDoubleAst, endDoubleAst + 2).clear();
                    textRange.getRange(m.start, m.start + 2).clear();

                    var styledRange = textRange.getRange(m.start, m.start + m.text.length);
                    styledRange.getTextStyle().setBold(true).setForegroundColor("#f67604");
                }
            } catch(e) {
                Logger.log("Erreur dans formatRichTextTech : " + e.message);
            }
        }

        var textMapping = {
            'TITRE_SLIDE_TECHNIQUE': props['TITRE_SLIDE_TECHNIQUE'] || "",
            'CRAWL_CONTENT_1': props['CRAWL_CONTENT_1'] || "",
            'CRAWL_CONTENT_2': props['CRAWL_CONTENT_2'] || "",
            'CRAWL_CONTENT_3': props['CRAWL_CONTENT_3'] || "",
            'INDEX_CONTENT_1': props['INDEX_CONTENT_1'] || "",
            'INDEX_CONTENT_2': props['INDEX_CONTENT_2'] || "",
            'INDEX_CONTENT_3': props['INDEX_CONTENT_3'] || "",
            'POS_CONTENT_1': props['POS_CONTENT_1'] || "",
            'POS_CONTENT_2': props['POS_CONTENT_2'] || "",
            'POS_CONTENT_3': props['POS_CONTENT_3'] || ""
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
                        formatRichTextTech(element);
                        return;
                    }

                    // Remplacement des icônes d'évaluation (BON, MOYEN, MAUVAIS, INCONNU)
                    var isIconPlaceholder = (
                        descRaw === "CRAWL_CHECK_1" || descRaw === "CRAWL_CHECK_2" || descRaw === "CRAWL_CHECK_3" ||
                        descRaw === "INDEX_CHECK_1" || descRaw === "INDEX_CHECK_2" || descRaw === "INDEX_CHECK_3" ||
                        descRaw === "POS_CHECK_1" || descRaw === "POS_CHECK_2" || descRaw === "POS_CHECK_3"
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
        var slideId = props['PA_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        var ICON_IDS = {
            "BON": "1lwxjX4LJWDoNYb19qco0VK93EH1V_aaQ",
            "MOYEN": "1l-eMhlZ4eXu2zxzH-_D_ZdRbIWB3X7VB",
            "MAUVAIS": "1WCVH1kIsBu5oEG_nWP9fQsGS5JZ5aGgI",
            "INCONNU": "1bi8wj96QvF9EetPHPEkVztTEwZf5H8tS"
        };

        function formatRichTextUX(shape) {
            try {
                var textRange = shape.getText();
                var textStr = textRange.asString();
                var regex = /\*([^*]+)\*/g;
                var matches = [];
                var match;

                while ((match = regex.exec(textStr)) !== null) {
                    matches.push({
                        start: match.index,
                        text: match[1],
                        length: match[0].length
                    });
                }
                
                for (var i = matches.length - 1; i >= 0; i--) {
                    var m = matches[i];
                    var endAst = m.start + m.text.length + 1;
                    
                    textRange.getRange(endAst, endAst + 1).clear();
                    textRange.getRange(m.start, m.start + 1).clear();

                    var styledRange = textRange.getRange(m.start, m.start + m.text.length);
                    styledRange.getTextStyle().setBold(true).setForegroundColor("#f67604");
                }
            } catch(e) {
                Logger.log("Erreur dans formatRichTextUX : " + e.message);
            }
        }

        var elementsMapping = {};
        for (var i = 1; i <= 6; i++) {
            elementsMapping['UX_ELEMENT_' + i] = props['UX_ELEMENT_' + i] || "";
        }
        
        var recoMapping = {
            'TITRE_SLIDE_UX': props['TITRE_SLIDE_UX'] || "",
            'UX_RECOMMANDATION_1': props['UX_RECOMMANDATION_1'] || "",
            'UX_RECOMMANDATION_2': props['UX_RECOMMANDATION_2'] || ""
        };

        var clientViewportId = props['UX_CLIENT_VIEWPORT_ID'];
        var compViewportId = props['UX_COMP_VIEWPORT_ID'];

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
                        formatRichTextUX(element);
                        return;
                    }

                    // Icônes d'évaluation
                    var isIconPlaceholder = descRaw.indexOf("UX_CLIENT_CHECK_") === 0 || descRaw.indexOf("UX_CONCURRENT_CHECK_") === 0;
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
                    if (descRaw === "PLACEHOLDER_UX_CLIENT" && clientViewportId && currentSlide) {
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

                    if (descRaw === "PLACEHOLDER_UX_CONCURRENT" && compViewportId && currentSlide) {
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

function exporterContenuSlides() {
    Logger.log("=== DÉBUT : exporterContenuSlides ===");
    try {
        var props = getDatabaseData();
        var slideId = props['PA_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        
        Logger.log("Ouverture de la présentation ID : " + slideId);
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        // ---------------------------------------------------------
        // FONCTIONS UTILITAIRES ENCAPSULÉES
        // ---------------------------------------------------------

        function formatRichTextContenu(shape) {
            Logger.log("  [Markdown] Traitement de la forme : " + shape.getDescription());
            try {
                var textRange = shape.getText();
                var textStr = textRange.asString();
                
                // 1. Gras standard (**)
                var regexGras = /\*\*([^*]+)\*\*/g;
                var matchesGras = [];
                var match;
                while ((match = regexGras.exec(textStr)) !== null) {
                    matchesGras.push({ start: match.index, text: match[1], length: match[0].length });
                }
                for (var i = matchesGras.length - 1; i >= 0; i--) {
                    var m = matchesGras[i];
                    textRange.getRange(m.start + m.text.length + 2, m.start + m.text.length + 4).clear();
                    textRange.getRange(m.start, m.start + 2).clear();
                    textRange.getRange(m.start, m.start + m.text.length).getTextStyle().setBold(true);
                }
                
                // 2. Gras orange (*)
                textStr = textRange.asString();
                var regexOrange = /\*([^*]+)\*/g;
                var matchesOrange = [];
                while ((match = regexOrange.exec(textStr)) !== null) {
                    matchesOrange.push({ start: match.index, text: match[1], length: match[0].length });
                }
                for (var j = matchesOrange.length - 1; j >= 0; j--) {
                    var mo = matchesOrange[j];
                    textRange.getRange(mo.start + mo.text.length + 1, mo.start + mo.text.length + 2).clear();
                    textRange.getRange(mo.start, mo.start + 1).clear();
                    textRange.getRange(mo.start, mo.start + mo.text.length).getTextStyle().setBold(true).setForegroundColor("#f67604");
                }
            } catch (e) {
                Logger.log("  [Erreur Markdown] " + e.message);
            }
        }

        function colorerScore1fr(shape, scoreStr) {
            Logger.log("  [Couleur 1.fr] Évaluation du score : " + scoreStr);
            try {
                var val = parseInt(scoreStr.replace(/[^0-9]/g, ''), 10);
                if (isNaN(val)) {
                    Logger.log("  [Couleur 1.fr] Score invalide, ignoré.");
                    return;
                }
                var color = "#ff0000"; // < 60%
                if (val >= 80) color = "#00b050";
                else if (val >= 60) color = "#f67604";
                
                Logger.log("  [Couleur 1.fr] Valeur extraite : " + val + " -> Couleur appliquée : " + color);
                shape.getFill().setSolidFill(color);
                shape.getText().getTextStyle().setForegroundColor("#ffffff").setBold(true);
            } catch (e) {
                Logger.log("  [Erreur Couleur 1.fr] " + e.message);
            }
        }

        function colorerScoreYTG(shape, scoreStr, cibleStr) {
            Logger.log("  [Couleur YTG] Évaluation du score : " + scoreStr + " (Cible : " + cibleStr + ")");
            try {
                var score = parseInt(scoreStr.replace(/[^0-9]/g, ''), 10);
                if (isNaN(score)) {
                    Logger.log("  [Couleur YTG] Score invalide, ignoré.");
                    return;
                }

                var min = 0, max = 999;
                var parts = cibleStr.split('-');
                if (parts.length === 2) {
                    min = parseInt(parts[0].replace(/[^0-9]/g, ''), 10);
                    max = parseInt(parts[1].replace(/[^0-9]/g, ''), 10);
                } else {
                    min = parseInt(cibleStr.replace(/[^0-9]/g, ''), 10);
                    max = min;
                }
                if (isNaN(min)) min = 0;
                if (isNaN(max)) max = 999;

                Logger.log("  [Couleur YTG] Fourchette cible parsée : [" + min + " - " + max + "]");

                var color = "#ff0000";
                if (score >= min && score <= max) {
                    color = "#00b050"; // Dans la fourchette
                    Logger.log("  [Couleur YTG] Résultat : Dans la cible -> Vert");
                } else if ((score >= min - 15 && score < min) || (score > max && score <= max + 15)) {
                    color = "#f67604"; // Tolérance 15 pts
                    Logger.log("  [Couleur YTG] Résultat : Dans la marge de tolérance (+/- 15) -> Orange");
                } else {
                    Logger.log("  [Couleur YTG] Résultat : Hors tolérance -> Rouge");
                }
                
                shape.getFill().setSolidFill(color);
                shape.getText().getTextStyle().setForegroundColor("#ffffff").setBold(true);
            } catch (e) {
                Logger.log("  [Erreur Couleur YTG] " + e.message);
            }
        }

        function insererEtRognerImage(slide, shape, fileId, description, cropId) {
            Logger.log("  [Image] Début traitement pour : " + description + " (ID Drive : " + fileId + ")");
            try {
                var finalFileId = fileId;
                if (cropId && cropId !== "") {
                    finalFileId = cropId;
                    Logger.log("  [Image] Utilisation du Crop de secours : " + cropId);
                }
                
                if (!finalFileId) throw new Error("ID Drive manquant.");
                var file = DriveApp.getFileById(finalFileId);
                var blob = file.getBlob().getAs(MimeType.JPEG);
                
                if (blob.getBytes().length === 0) {
                    throw new Error("Le fichier est vide ou corrompu.");
                }

                Logger.log("  [Image] Remplacement natif avec option crop=true (Fill Box)...");
                // Le paramètre "true" force l'image à remplir totalement le placeholder sans le déformer.
                // Il applique un "Center Crop" automatique. Cela préserve le Z-index, le groupe, et les styles.
                var img = shape.replaceWithImage(blob, true);
                
                // On lit les valeurs de rognage calculées automatiquement par l'API
                var currentCropTop = img.getCropTop();
                var currentCropBottom = img.getCropBottom();
                var totalVerticalCrop = currentCropTop + currentCropBottom;
                
                // On décale le "fenêtrage" vers le haut : on met le rognage haut à 0 
                // et on reporte la totalité du rognage sur le bas.
                img.setCropTop(0);
                img.setCropBottom(totalVerticalCrop);

                img.setDescription(description);
                Logger.log("  [Image] Rognage ajusté vers le haut (Top: 0, Bottom: " + totalVerticalCrop.toFixed(4) + ").");
            } catch (e) {
                Logger.log("  [Erreur Image] Échec pour " + description + " : " + e.message);
            }
        }

        // ---------------------------------------------------------
        // MAPPING ET BOUCLE PRINCIPALE
        // ---------------------------------------------------------

        var textMapping = {
            'TITRE_SLIDE_CONTENU_CLIENT': props['TITRE_SLIDE_CONTENU_CLIENT'] || "",
            'TITRE_SLIDE_CONTENU_CONCURRENT': props['TITRE_SLIDE_CONTENU_CONCURRENT'] || "",
            'CONTENU_STRUCTURE_CLIENT': props['CONTENU_STRUCTURE_CLIENT'] || "",
            'CONTENU_YTG_CLIENT': props['CONTENU_YTG_CLIENT'] || "",
            'CONTENU_1FR_CLIENT': props['CONTENU_1FR_CLIENT'] || "",
            'CONTENU_STRUCTURE_CONCURRENT': props['CONTENU_STRUCTURE_CONCURRENT'] || "",
            'CONTENU_YTG_CONCURRENT': props['CONTENU_YTG_CONCURRENT'] || "",
            'CONTENU_1FR_CONCURRENT': props['CONTENU_1FR_CONCURRENT'] || ""
        };

        var cibleYTG = props['CONTENU_YTG_CIBLE'] || "";
        var clientFullId = props['UX_CLIENT_FULL_ID'];
        var compFullId = props['UX_COMP_FULL_ID'];
        var clientCropId = props['UX_CLIENT_CROP_ID'];
        var compCropId = props['UX_COMP_CROP_ID'];

        Logger.log("Début du parcours récursif des slides...");

        slides.forEach(function(slide, slideIndex) {
            
            function processElement(element) {
                var type = element.getPageElementType();
                
                if (type === SlidesApp.PageElementType.GROUP) {
                    element.asGroup().getChildren().forEach(processElement);
                } 
                else if (type === SlidesApp.PageElementType.SHAPE) {
                    var shape = element.asShape();
                    var descRaw = shape.getDescription() || "";

                    if (descRaw !== "") {
                        // 1. Textes classiques avec Markdown
                        if (textMapping[descRaw] !== undefined) {
                            Logger.log("Remplacement du texte pour : " + descRaw);
                            shape.getText().setText(String(textMapping[descRaw]));
                            formatRichTextContenu(shape);
                            return;
                        }

                        // 2. Cible YTG (texte simple)
                        if (descRaw === 'CONTENU_YTG_CIBLE') {
                            Logger.log("Remplacement de la cible YTG");
                            shape.getText().setText(String(cibleYTG));
                            return;
                        }

                        // 3. Scores 1.fr
                        if (descRaw === 'CONTENU_1FR_SCORE_CLIENT') {
                            var score = props['CONTENU_1FR_SCORE_CLIENT'] || "";
                            shape.getText().setText(String(score));
                            colorerScore1fr(shape, score);
                            return;
                        }
                        if (descRaw === 'CONTENU_1FR_SCORE_CONCURRENT') {
                            var scoreComp = props['CONTENU_1FR_SCORE_CONCURRENT'] || "";
                            shape.getText().setText(String(scoreComp));
                            colorerScore1fr(shape, scoreComp);
                            return;
                        }

                        // 4. Scores YTG
                        if (descRaw === 'CONTENU_YTG_SCORE_CLIENT') {
                            var scoreYtgCli = props['CONTENU_YTG_SCORE_CLIENT'] || "";
                            shape.getText().setText(String(scoreYtgCli));
                            colorerScoreYTG(shape, scoreYtgCli, cibleYTG);
                            return;
                        }
                        if (descRaw === 'CONTENU_YTG_SCORE_CONCURRENT') {
                            var scoreYtgComp = props['CONTENU_YTG_SCORE_CONCURRENT'] || "";
                            shape.getText().setText(String(scoreYtgComp));
                            colorerScoreYTG(shape, scoreYtgComp, cibleYTG);
                            return;
                        }

                        // 5. Images Full Page
                        if (descRaw === 'PLACEHOLDER_CONTENU_CLIENT') {
                            insererEtRognerImage(slide, shape, clientFullId, descRaw, clientCropId);
                            return;
                        }
                        if (descRaw === 'PLACEHOLDER_CONTENU_CONCURRENT') {
                            insererEtRognerImage(slide, shape, compFullId, descRaw, compCropId);
                            return;
                        }
                    }
                }
            }

            slide.getPageElements().forEach(processElement);
        });

        Logger.log("=== FIN : exporterContenuSlides (Succès) ===");
        return { success: true, url: presentation.getUrl() };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT CONTENU : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterEditorialSlides(concurrenceDataEdito, titreSlide, donneesBlog) {
    try {
        Logger.log("=== DÉBUT : exporterEditorialSlides ===");
        var props = getDatabaseData();
        var slideId = props['PA_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        function safeNum(val) {
            return (val !== null && val !== undefined && !isNaN(val)) ? Math.round(val).toLocaleString('fr-FR') : "-";
        }

        var mappingComp = {};
        
        var titreThematique = props['TITRE_SLIDE_THEMATIQUE_EDITO'] || "";
        mappingComp['TITRE_SLIDE_CONCURRENCE_EDITO'] = titreSlide || "";
        mappingComp['TITRE_SLIDE_THEMATIQUE_EDITO'] = titreThematique;
        
        if (concurrenceDataEdito) {
            if (concurrenceDataEdito.client) {
                mappingComp['NOM_CLIENT_EDITO'] = concurrenceDataEdito.client.name;
                mappingComp['VALEUR_TOP10_CLIENT_EDITO'] = safeNum(concurrenceDataEdito.client.top10);
                mappingComp['VALEUR_PAGES_CLIENT_EDITO'] = safeNum(concurrenceDataEdito.client.pages);
                mappingComp['BLOG_CLIENT_EDITO'] = donneesBlog.client || "Non";
            }
            if (concurrenceDataEdito.leader) {
                mappingComp['NOM_LEADER_EDITO'] = concurrenceDataEdito.leader.name;
                mappingComp['VALEUR_TOP10_LEADER_EDITO'] = safeNum(concurrenceDataEdito.leader.top10);
                mappingComp['VALEUR_PAGES_LEADER_EDITO'] = safeNum(concurrenceDataEdito.leader.pages);
                mappingComp['BLOG_LEADER_EDITO'] = donneesBlog.leader || "Non";
            }
            for (var c = 1; c <= 4; c++) {
                var comp = concurrenceDataEdito.comps && concurrenceDataEdito.comps[c-1] ? concurrenceDataEdito.comps[c-1] : null;
                if (comp) {
                    mappingComp['NOM_COMP' + c + '_EDITO'] = comp.name;
                    mappingComp['VALEUR_TOP10_COMP' + c + '_EDITO'] = safeNum(comp.top10);
                    mappingComp['VALEUR_PAGES_COMP' + c + '_EDITO'] = safeNum(comp.pages);
                    mappingComp['BLOG_COMP' + c + '_EDITO'] = donneesBlog['comp' + c] || "Non";
                }
            }
            
            // Traitement des pistes (Thématiques)
            var selectionJson = props['ANALYSE_SELECTION'];
            var selection = selectionJson ? JSON.parse(selectionJson) : [];
            var pistesEdito = [];
            try {
                var diag = genererDiagnostic(selection);
                if (diag && diag.pistesEdito) pistesEdito = diag.pistesEdito;
            } catch (e) {
                Logger.log("Erreur diagnostic dans exporterEditorialSlides : " + e.message);
            }
            
            for (var i = 1; i <= 3; i++) {
                var piste = pistesEdito[i - 1];
                if (piste) {
                    mappingComp['THEMATIQUE_EDITO_' + i] = piste.tsKey || "";
                    mappingComp['NOM_COMP' + i + '_EDITO_CONTENU'] = piste.entite || "";
                    mappingComp['URL_EDITO_CONTENU_' + i] = piste.url || "";
                    mappingComp['NOM_CONTENU_' + i] = props['NOM_CONTENU_' + i] || "";
                    
                    var kwStr = piste.kwTop100 + " mots-clés positionnés en top 100\n\n" + piste.kwTop10 + " mots-clés positionnés en top 10, exemples :\n";
                    if (piste.kws && piste.kws.length > 0) {
                        for (var k = 0; k < piste.kws.length; k++) {
                            var kwItem = piste.kws[k];
                            kwStr += "• « " + kwItem.kw + " », recherché " + safeNum(kwItem.vol) + " fois par mois, en position " + kwItem.pos + "\n";
                        }
                    }
                    // Retirer le dernier saut de ligne
                    kwStr = kwStr.replace(/\n$/, "");
                    mappingComp['DATA_TOP10_CONTENU_' + i] = kwStr;
                } else {
                    mappingComp['THEMATIQUE_EDITO_' + i] = "-";
                    mappingComp['NOM_COMP' + i + '_EDITO_CONTENU'] = "-";
                    mappingComp['URL_EDITO_CONTENU_' + i] = "-";
                    mappingComp['NOM_CONTENU_' + i] = "-";
                    mappingComp['DATA_TOP10_CONTENU_' + i] = "-";
                }
            }
        }

        // Sauvegarde des propriétés
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

        Logger.log("Parcours minutieux des slides pour Performance Éditoriale");

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();

            shapes.forEach(function(shape) {
                var descRaw = shape.getDescription() || "";

                if (mappingComp[descRaw] !== undefined) {
                    shape.getText().setText(mappingComp[descRaw].toString());
                }

                // Placeholders d'images concurrentes pour l'éditorial
                if (descRaw.indexOf("PLACEHOLDER_LOGO_") === 0 && descRaw.indexOf("_EDITO") !== -1) {
                    var imgUrl = null;
                    
                    if (descRaw.indexOf("_EDITO_CONTENU") !== -1) {
                        var mc = descRaw.match(/PLACEHOLDER_LOGO_COMP(\d+)_EDITO_CONTENU/);
                        if (mc && mc[1]) {
                            var iContenu = parseInt(mc[1], 10);
                            var domainUrl = mappingComp['URL_EDITO_CONTENU_' + iContenu];
                            if (domainUrl && domainUrl !== "-") {
                                var cleanDomain = domainUrl.replace(/^(?:https?:\/\/)?(?:www\.)?/i, "").split('/')[0];
                                imgUrl = 'https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&url=https://' + cleanDomain + '&size=128';
                            }
                        }
                    } else {
                        if (descRaw === "PLACEHOLDER_LOGO_CLIENT_EDITO" && concurrenceDataEdito.client && concurrenceDataEdito.client.logoUrl) {
                            imgUrl = concurrenceDataEdito.client.logoUrl;
                        } else if (descRaw === "PLACEHOLDER_LOGO_LEADER_EDITO" && concurrenceDataEdito.leader && concurrenceDataEdito.leader.logoUrl) {
                            imgUrl = concurrenceDataEdito.leader.logoUrl;
                        } else {
                            var m = descRaw.match(/PLACEHOLDER_LOGO_COMP(\d+)_EDITO/);
                            if (m && m[1]) {
                                var idxComp = parseInt(m[1]) - 1;
                                if (concurrenceDataEdito.comps && concurrenceDataEdito.comps[idxComp] && concurrenceDataEdito.comps[idxComp].logoUrl) {
                                    imgUrl = concurrenceDataEdito.comps[idxComp].logoUrl;
                                }
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
            });
        });

        Logger.log("=== FIN : exporterEditorialSlides ===");
        return { success: true, url: presentation.getUrl() };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT EDITORIAL : " + e.message);
        return { success: false, error: e.message };
    }
}