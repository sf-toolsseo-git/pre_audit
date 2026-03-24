function exporterSlideBesoinSolution(texteBesoin, texteSolution) {
    try {
        Logger.log("=== DÉBUT : exporterSlideBesoinSolution ===");
        var props = PropertiesService.getScriptProperties().getProperties();
        var slideId = props['PA_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();
        
        var tagsTrouves = 0;
        
        // Formatage : on ajoute un saut de ligne vide entre chaque puce pour aérer dans Slides
        var slideTexteBesoin = texteBesoin.replace(/\n/g, '\n\n');
        var slideTexteSolution = texteSolution.replace(/\n/g, '\n\n');
        
        Logger.log("Recherche des tags TAG_SLIDE_BESOIN et TAG_SLIDE_SOLUTION");

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();
            
            shapes.forEach(function(shape) {
                var titleRaw = shape.getTitle() || "";
                var descRaw = shape.getDescription() || "";

                var targetKey = null;
                // Détection via le texte alternatif en majuscule
                if (titleRaw === "TAG_SLIDE_BESOIN" || descRaw === "TAG_SLIDE_BESOIN") {
                    targetKey = "besoin";
                } else if (titleRaw === "TAG_SLIDE_SOLUTION" || descRaw === "TAG_SLIDE_SOLUTION") {
                    targetKey = "solution";
                }

                if (targetKey === "besoin") {
                    shape.getText().setText(slideTexteBesoin);
                    tagsTrouves++;
                } else if (targetKey === "solution") {
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
        var props = PropertiesService.getScriptProperties().getProperties();
        var slideId = props['PA_SLIDE_ID'];
        
        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        // Helper pour appliquer le style "Open Sans", taille 18, et traiter les termes en gras orange
        function formatRichText(shape, textWithStars) {
            var textRange = shape.getText();
            var cleanText = textWithStars.replace(/\*/g, ""); 
            
            textRange.setText(cleanText);
            textRange.getTextStyle().setFontFamily("Open Sans").setFontSize(16).setForegroundColor("#000000").setBold(false);

            var match;
            var regex = /\*([^*]+)\*/g;
            var offset = 0;
            while ((match = regex.exec(textWithStars)) !== null) {
                var wordLength = match[1].length;
                var startIndex = match.index - offset;
                var endIndex = startIndex + wordLength;
                
                var targetRange = textRange.getRange(startIndex, endIndex);
                targetRange.getTextStyle().setBold(true).setForegroundColor("#f67604");
                offset += 2; 
            }
        }

        Logger.log("Parcours des slides pour l'analyse Semrush en cours");

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();
            
            shapes.forEach(function(shape) {
                
                var shapeText = shape.getText().asString().trim();
                var titleRaw = shape.getTitle() || "";
                var descRaw = shape.getDescription() || "";

                // Remplacement du titre brut
                if (shapeText === "TITRE_SLIDE_SEMRUSH" || titleRaw === "TITRE_SLIDE_SEMRUSH" || descRaw === "TITRE_SLIDE_SEMRUSH") {
                    shape.getText().setText(titre);
                }
                
                // Remplacement des analyses avec formatage enrichi
                if (shapeText === "ANALYSE_SEMRUSH_MOT_CLE" || titleRaw === "ANALYSE_SEMRUSH_MOT_CLE" || descRaw === "ANALYSE_SEMRUSH_MOT_CLE") {
                    formatRichText(shape, texteKw);
                }
                
                if (shapeText === "ANALYSE_SEMRUSH_TRAFIC" || titleRaw === "ANALYSE_SEMRUSH_TRAFIC" || descRaw === "ANALYSE_SEMRUSH_TRAFIC") {
                    formatRichText(shape, texteTrafic);
                }

                // Remplacement des images avec les nouveaux placeholders
                if (titleRaw === "PLACEHOLDER_ANALYSE_SEMRUSH_MOT_CLE" || descRaw === "PLACEHOLDER_ANALYSE_SEMRUSH_MOT_CLE" || shapeText === "PLACEHOLDER_ANALYSE_SEMRUSH_MOT_CLE") {
                    Logger.log("Remplacement image mots-clés");
                    var blobKw = Utilities.newBlob(Utilities.base64Decode(imgKwB64), imgKwMime, "kw.png");
                    var newImageKw = slide.insertImage(blobKw, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                    
                    newImageKw.setTitle(titleRaw);
                    newImageKw.setDescription(descRaw);
                    shape.remove();
                }
                
                if (titleRaw === "PLACEHOLDER_ANALYSE_SEMRUSH_TRAFIC" || descRaw === "PLACEHOLDER_ANALYSE_SEMRUSH_TRAFIC" || shapeText === "PLACEHOLDER_ANALYSE_SEMRUSH_TRAFIC") {
                    Logger.log("Remplacement image trafic");
                    var blobTrafic = Utilities.newBlob(Utilities.base64Decode(imgTraficB64), imgTraficMime, "trafic.png");
                    var newImageTrafic = slide.insertImage(blobTrafic, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                    
                    newImageTrafic.setTitle(titleRaw);
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
        var props = PropertiesService.getScriptProperties().getProperties();
        var slideId = props['PA_SLIDE_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        // 1. Extraction des données
        var clientKpi = diagnosticData.kpis.find(function(k) { return k.isClient; });
        if (!clientKpi) throw new Error("Données client introuvables dans le diagnostic.");

        var intentStats = diagnosticData.intentStats;
        var totalTop10 = clientKpi.top10;
        // 2. Calcul des pourcentages (Parts de l'intention dans le top 10 global du client)
        var transacPctDec = totalTop10 > 0 ? (intentStats.transac.top10 / totalTop10) : 0;
        var infoPctDec = totalTop10 > 0 ? (intentStats.info.top10 / totalTop10) : 0;
        
        // 3. Préparation des tableaux triés (Thèmes et Segments)
        Logger.log("Préparation des tableaux triés pour l'injection...");
        var topThemes = diagnosticData.themeStats ? diagnosticData.themeStats.slice().sort(function(a, b) { return b.TEC - a.TEC || b.top10 - a.top10 || b.top3 - a.top3; }).slice(0, 3) : [];
        var flopThemes = diagnosticData.themeStats ? diagnosticData.themeStats.slice().sort(function(a, b) { return b.DDT - a.DDT; }).slice(0, 3) : [];
        var acquis = diagnosticData.acquis ? diagnosticData.acquis.slice(0, 5) : [];
        var gains = diagnosticData.gains ? diagnosticData.gains.slice(0, 5) : [];
        var pertes = diagnosticData.pertes ? diagnosticData.pertes.slice(0, 5) : [];
        var territoires = diagnosticData.territoires ? diagnosticData.territoires.slice(0, 5) : [];
        
        // Fonctions utilitaires pour éviter les erreurs "undefined"
        function safeNum(val) {
            return (val !== null && val !== undefined && !isNaN(val)) ? Math.round(val).toLocaleString('fr-FR') : "-";
        }
        function safePos(val) {
            return (val !== null && val !== undefined && !isNaN(val)) ? Number(val).toLocaleString('fr-FR') : "-";
        }

        // 4. Mapping textuel & dictionnaire de remplacement
        // Mapping paysage concurrentiel
        var mappingComp = {};
        if (concurrenceData) {
            mappingComp['TITRE_SLIDE_CONCURRENCE'] = "L'environnement concurrentiel de " + (concurrenceData.client ? concurrenceData.client.name : "");
            // Client
            if (concurrenceData.client) {
                mappingComp['NOM_CLIENT'] = concurrenceData.client.name;
                mappingComp['VALEUR_TOP10_CLIENT'] = safeNum(concurrenceData.client.top10);
                mappingComp['VALEUR_PAGES_CLIENT'] = safeNum(concurrenceData.client.pages);
            }
            // Leader
            if (concurrenceData.leader) {
                mappingComp['NOM_LEADER'] = concurrenceData.leader.name;
                mappingComp['VALEUR_TOP10_LEADER'] = safeNum(concurrenceData.leader.top10);
                mappingComp['VALEUR_PAGES_LEADER'] = safeNum(concurrenceData.leader.pages);
            } else {
                mappingComp['NOM_LEADER'] = "";
                mappingComp['VALEUR_TOP10_LEADER'] = "";
                mappingComp['VALEUR_PAGES_LEADER'] = "";
            }
            // Concurrents
            for (var c = 1; c <= 4; c++) {
                var comp = concurrenceData.comps && concurrenceData.comps[c-1] ? concurrenceData.comps[c-1] : null;
                if (comp) {
                    mappingComp['NOM_COMP' + c] = comp.name;
                    mappingComp['VALEUR_TOP10_COMP' + c] = safeNum(comp.top10);
                    mappingComp['VALEUR_PAGES_COMP' + c] = safeNum(comp.pages);
                } else {
                    mappingComp['NOM_COMP' + c] = "";
                    mappingComp['VALEUR_TOP10_COMP' + c] = "";
                    mappingComp['VALEUR_PAGES_COMP' + c] = "";
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
        // Thèmes Top (1 à 3)
        for (var i = 1; i <= 3; i++) {
            var thm = topThemes[i - 1];
            replaceDict["{{top_thm_client_" + i + "}}"] = thm ? thm.name : "-";
            replaceDict["{{top_thm_client_top10_" + i + "}}"] = thm ? safeNum(thm.top10) : "-";
            replaceDict["{{top_thm_client_tec_" + i + "}}"] = thm ? safeNum(thm.TEC) : "-";
            replaceDict["{{top_thm_client_tpm_" + i + "}}"] = thm ? safeNum(thm.TPM) : "-";
            replaceDict["{{top_thm_client_ddt_" + i + "}}"] = thm ? safeNum(thm.DDT) : "-";
        }

        // Thèmes Flop (1 à 3)
        for (var i = 1; i <= 3; i++) {
            var thm = flopThemes[i - 1];
            replaceDict["{{flop_thm_client_" + i + "}}"] = thm ? thm.name : "-";
            replaceDict["{{flop_thm_client_flop10_" + i + "}}"] = thm ? safeNum(thm.top10) : "-";
            replaceDict["{{flop_thm_client_tec_" + i + "}}"] = thm ? safeNum(thm.TEC) : "-";
            replaceDict["{{flop_thm_client_tpm_" + i + "}}"] = thm ? safeNum(thm.TPM) : "-";
            replaceDict["{{flop_thm_client_ddt_" + i + "}}"] = thm ? safeNum(thm.DDT) : "-";
        }

        // Segments (1 à 5)
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

        // Helper pour découper l'analyse IA en 3 blocs distincts
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

        // ==========================================
        // SAUVEGARDE EN BASE DES VARIABLES DYNAMIQUES
        // ==========================================
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

        // Forçage de la valeur "IMAGE" pour tous les placeholders de logos afin qu'ils apparaissent correctement dans CONFIG
        propsToSave["PLACEHOLDER_LOGO_CLIENT"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_LEADER"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP1"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP2"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP3"] = "IMAGE";
        propsToSave["PLACEHOLDER_LOGO_COMP4"] = "IMAGE";

        PropertiesService.getScriptProperties().setProperties(propsToSave);
        syncPropertiesToConfigSheet();
        // ==========================================

        Logger.log("Remplacement massif des tags de tableaux en cours...");
        for (var key in replaceDict) {
            presentation.replaceAllText(key, String(replaceDict[key]));
        }
        Logger.log("Remplacement massif terminé.");
        
        // Fonction utilitaire de formatage pour l'IA
        function formatRichText(shape, textWithStars) {
            if (!textWithStars) return;
            var textRange = shape.getText();
            var cleanText = textWithStars.replace(/\*/g, "");
            
            textRange.setText(cleanText);
            textRange.getTextStyle().setFontFamily("Open Sans").setFontSize(18).setForegroundColor("#000000").setBold(false);
            var match;
            var regex = /\*([^*]+)\*/g;
            var offset = 0;
            while ((match = regex.exec(textWithStars)) !== null) {
                var wordLength = match[1].length;
                var startIndex = match.index - offset;
                var endIndex = startIndex + wordLength;
                
                var targetRange = textRange.getRange(startIndex, endIndex);
                targetRange.getTextStyle().setBold(true).setForegroundColor("#f67604");
                offset += 2;
            }
        }

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();

            // Étape A : Remplacement des textes classiques et IA
            shapes.forEach(function(shape) {
                var shapeText = shape.getText().asString().trim();
                var titleRaw = shape.getTitle() || "";
                var descRaw = shape.getDescription() || "";

                // 1. Textes basiques (KPI client & intent)
                var targetKey = mapping[shapeText] ? shapeText : (mapping[titleRaw] ? titleRaw : (mapping[descRaw] ? descRaw : null));

                if (targetKey) {
                    shape.getText().setText(mapping[targetKey].toString());
                }

                var targetCompKey = mappingComp && mappingComp[shapeText] !== undefined ? shapeText : (mappingComp && mappingComp[titleRaw] !== undefined ? titleRaw : (mappingComp && mappingComp[descRaw] !== undefined ? descRaw : null));
                if (targetCompKey) {
                    shape.getText().setText(mappingComp[targetCompKey].toString());
                }

                // Traitement des favicons du paysage concurrentiel
                if (concurrenceData && (titleRaw.indexOf("PLACEHOLDER_LOGO_") === 0 || descRaw.indexOf("PLACEHOLDER_LOGO_") === 0 || shapeText.indexOf("PLACEHOLDER_LOGO_") === 0)) {
                    var tagParts = titleRaw.indexOf("PLACEHOLDER_LOGO_") === 0 ? titleRaw : (descRaw.indexOf("PLACEHOLDER_LOGO_") === 0 ? descRaw : shapeText);
                    var imgUrl = null;
                    
                    if (tagParts === "PLACEHOLDER_LOGO_CLIENT" && concurrenceData.client && concurrenceData.client.logoUrl) {
                        imgUrl = concurrenceData.client.logoUrl;
                    } else if (tagParts === "PLACEHOLDER_LOGO_LEADER" && concurrenceData.leader && concurrenceData.leader.logoUrl) {
                        imgUrl = concurrenceData.leader.logoUrl;
                    } else {
                        var m = tagParts.match(/PLACEHOLDER_LOGO_COMP(\d+)/);
                        if (m && m[1]) {
                            var idx = parseInt(m[1]) - 1;
                            if (concurrenceData.comps && concurrenceData.comps[idx] && concurrenceData.comps[idx].logoUrl) {
                                imgUrl = concurrenceData.comps[idx].logoUrl;
                            }
                        }
                    }

                    if (imgUrl) {
                        try {
                            Logger.log("Téléchargement de l'image : " + imgUrl);
                            var response = UrlFetchApp.fetch(imgUrl, { muteHttpExceptions: true });
                            if (response.getResponseCode() === 200) {
                                var blob = response.getBlob();
                                var newImg = slide.insertImage(blob, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                                newImg.setTitle(titleRaw);
                                newImg.setDescription(descRaw);
                            } else {
                                Logger.log("Image ignorée (code " + response.getResponseCode() + ") : " + imgUrl);
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

                // 2. Titres IA
                if (shapeText === "TITRE_SLIDE_THEMATIQUETOP_CLIENT" || titleRaw === "TITRE_SLIDE_THEMATIQUETOP_CLIENT" || descRaw === "TITRE_SLIDE_THEMATIQUETOP_CLIENT") {
                    if (iaData && iaData.titreTopThematiques) shape.getText().setText(iaData.titreTopThematiques);
                }
                if (shapeText === "TITRE_SLIDE_THEMATIQUEFLOP_CLIENT" || titleRaw === "TITRE_SLIDE_THEMATIQUEFLOP_CLIENT" || descRaw === "TITRE_SLIDE_THEMATIQUEFLOP_CLIENT") {
                    if (iaData && iaData.titreFlopThematiques) shape.getText().setText(iaData.titreFlopThematiques);
                }
                if (shapeText === "TITRE_SLIDE_MCTOP_CLIENT" || titleRaw === "TITRE_SLIDE_MCTOP_CLIENT" || descRaw === "TITRE_SLIDE_MCTOP_CLIENT") {
                    if (iaData && iaData.titreTopSegments) shape.getText().setText(iaData.titreTopSegments);
                }
                if (shapeText === "TITRE_SLIDE_MCFLOP_CLIENT" || titleRaw === "TITRE_SLIDE_MCFLOP_CLIENT" || descRaw === "TITRE_SLIDE_MCFLOP_CLIENT") {
                    if (iaData && iaData.titreFlopSegments) shape.getText().setText(iaData.titreFlopSegments);
                }

                // 3. Analyses IA découpées en blocs (1, 2, 3)
                for (var idx = 1; idx <= 3; idx++) {
                    var tagTopThem = "ANALYSE_THEMATIQUETOP_CLIENT_" + idx;
                    if (shapeText === tagTopThem || titleRaw === tagTopThem || descRaw === tagTopThem) {
                        formatRichText(shape, topThemParts[idx-1]);
                    }

                    var tagFlopThem = "ANALYSE_THEMATIQUEFLOP_CLIENT_" + idx;
                    if (shapeText === tagFlopThem || titleRaw === tagFlopThem || descRaw === tagFlopThem) {
                        formatRichText(shape, flopThemParts[idx-1]);
                    }

                    var tagTopSeg = "ANALYSE_MCTOP_CLIENT_" + idx;
                    if (shapeText === tagTopSeg || titleRaw === tagTopSeg || descRaw === tagTopSeg) {
                        formatRichText(shape, topSegParts[idx-1]);
                    }

                    var tagFlopSeg = "ANALYSE_MCFLOP_CLIENT_" + idx;
                    if (shapeText === tagFlopSeg || titleRaw === tagFlopSeg || descRaw === tagFlopSeg) {
                        formatRichText(shape, flopSegParts[idx-1]);
                    }
                }
            });
            
            // Étape B : Traitement des jauges dynamiques (Conservation de l'arrondi)
            var shapesForGauges = slide.getShapes();
            shapesForGauges.forEach(function(shape) {
                var shapeText = shape.getText().asString().trim();
                var titleRaw = shape.getTitle() || "";
                var descRaw = shape.getDescription() || "";
                
                var isTransacGauge = (titleRaw === "JAUGE_TRANSAC_TOP10" || descRaw === "JAUGE_TRANSAC_TOP10" || shapeText === "JAUGE_TRANSAC_TOP10");
                var isInfoGauge = (titleRaw === "JAUGE_INFO_TOP10" || descRaw === "JAUGE_INFO_TOP10" || shapeText === "JAUGE_INFO_TOP10");

                var targetGauge = isTransacGauge ? "transac" : (isInfoGauge ? "info" : null);

                if (targetGauge) {
                    var pct = (targetGauge === "transac") ? transacPctDec : infoPctDec;

                    var left = shape.getLeft();
                    var top = shape.getTop();
                    var width = shape.getWidth();

                    // 1. Dessiner la jauge verte en DUPLIQUANT la forme originale
                    if (pct > 0) {
                        var fgShape = slide.insertShape(shape);
                        var fillWidth = Math.max(shape.getHeight(), width * pct);
                        fgShape.setWidth(fillWidth);
                        fgShape.setLeft(left);
                        fgShape.setTop(top);
                        
                        // Réalignement strict
                        fgShape.getFill().setSolidFill("#00b050");
                        fgShape.getBorder().setTransparent();
                        fgShape.getText().clear();
                        fgShape.setTitle("");
                        fgShape.setDescription("");
                    }

                    // 2. Transformer la forme originale en fond gris
                    shape.getFill().setSolidFill("#f1f3f4");
                    shape.getBorder().setTransparent();
                    shape.getText().clear();
                    shape.setTitle("");
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