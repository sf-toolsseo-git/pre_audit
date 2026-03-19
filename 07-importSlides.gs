function exporterSlideBesoinSolution(texteBesoin, texteSolution) {
    try {
        Logger.log("=== DÉBUT EXPORT SLIDE BESOIN / SOLUTION ===");
        var props = PropertiesService.getScriptProperties().getProperties();
        var slideId = props['SLIDE_PRE_AUDIT_ID'];

        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();
        
        var tagsTrouves = 0;
        slides.forEach(function(slide) {
            var shapes = slide.getShapes();
            
            shapes.forEach(function(shape) {
                var titleRaw = shape.getTitle() || "";
                var descRaw = shape.getDescription() || "";

                var targetKey = null;
                // Détection via le texte alternatif (titre ou description) uniquement
                if (titleRaw === "tag_slide_besoin" || descRaw === "tag_slide_besoin") {
                    targetKey = "besoin";
                } else if (titleRaw === "tag_slide_solution" || descRaw === "tag_slide_solution") {
                    targetKey = "solution";
                }

                if (targetKey === "besoin") {
                    shape.getText().setText(texteBesoin);
                    tagsTrouves++;
                } else if (targetKey === "solution") {
                    shape.getText().setText(texteSolution);
                    tagsTrouves++;
                }
            });
        });
        
        Logger.log("=== FIN EXPORT SLIDE BESOIN / SOLUTION ===");
        
        if (tagsTrouves === 0) {
            return { success: false, error: "Les tags 'tag_slide_besoin' et 'tag_slide_solution' n'ont pas été trouvés dans le texte alternatif de la présentation." };
        }

        return { success: true, url: presentation.getUrl() };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT BESOIN/SOLUTION : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterAnalyseSemrushSlide(titre, texteKw, texteTrafic, imgKwB64, imgKwMime, imgTraficB64, imgTraficMime) {
    try {
        Logger.log("=== DÉBUT EXPORT SLIDE SEMRUSH ===");
        var props = PropertiesService.getScriptProperties().getProperties();
        var slideId = props['SLIDE_PRE_AUDIT_ID'];
        
        if (!slideId) throw new Error("L'ID du Google Slides n'est pas configuré.");
        var presentation = SlidesApp.openById(slideId);
        var slides = presentation.getSlides();

        // Helper pour appliquer le style "Open Sans", taille 18, et traiter les *termes en gras orange*
        function formatRichText(shape, textWithStars) {
            var textRange = shape.getText();
            var cleanText = textWithStars.replace(/\*/g, ""); // Texte sans les astérisques
            
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
                
                offset += 2; // On compense les 2 astérisques retirés pour les calculs des index suivants
            }
        }

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();
            
            shapes.forEach(function(shape) {
                var shapeText = shape.getText().asString().trim();
                var titleRaw = shape.getTitle() || "";
                var descRaw = shape.getDescription() || "";

                // Remplacement du titre brut
                if (shapeText === "titre_slide_semrush" || titleRaw === "titre_slide_semrush" || descRaw === "titre_slide_semrush") {
                    shape.getText().setText(titre);
                }
                
                // Remplacement des analyses avec formatage enrichi
                if (shapeText === "analyse_semrush_mot_cle" || titleRaw === "analyse_semrush_mot_cle" || descRaw === "analyse_semrush_mot_cle") {
                    formatRichText(shape, texteKw);
                }
                
                if (shapeText === "analyse_semrush_trafic" || titleRaw === "analyse_semrush_trafic" || descRaw === "analyse_semrush_trafic") {
                    formatRichText(shape, texteTrafic);
                }

                // Remplacement des images avec conservation du texte alternatif
                if (titleRaw === "placeholder_img_kw" || descRaw === "placeholder_img_kw" || shapeText === "placeholder_img_kw") {
                    var blobKw = Utilities.newBlob(Utilities.base64Decode(imgKwB64), imgKwMime, "kw.png");
                    var newImageKw = slide.insertImage(blobKw, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                    
                    // Conservation du texte alternatif
                    newImageKw.setTitle(titleRaw);
                    newImageKw.setDescription(descRaw);
                    
                    shape.remove();
                }
                
                if (titleRaw === "placeholder_img_trafic" || descRaw === "placeholder_img_trafic" || shapeText === "placeholder_img_trafic") {
                    var blobTrafic = Utilities.newBlob(Utilities.base64Decode(imgTraficB64), imgTraficMime, "trafic.png");
                    var newImageTrafic = slide.insertImage(blobTrafic, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
                    
                    // Conservation du texte alternatif
                    newImageTrafic.setTitle(titleRaw);
                    newImageTrafic.setDescription(descRaw);
                    
                    shape.remove();
                }
            });
        });

        Logger.log("=== FIN EXPORT SLIDE SEMRUSH ===");
        return { success: true, url: presentation.getUrl() };

    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT SLIDE SEMRUSH : " + e.message);
        return { success: false, error: e.message };
    }
}

function exporterPerformanceGlobalSlides(diagnosticData) {
    try {
        Logger.log("=== DÉBUT EXPORT SLIDE GLOBAL & INTENTIONS ===");
        var props = PropertiesService.getScriptProperties().getProperties();
        var slideId = props['SLIDE_PRE_AUDIT_ID'];

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

        // 3. Mapping textuel
        var mapping = {
            'mot_cle_client_global': clientKpi.posAll.toLocaleString('fr-FR'),
            'mot_cle_client_top3': clientKpi.top3.toLocaleString('fr-FR'),
            'mot_cle_client_top10': clientKpi.top10.toLocaleString('fr-FR'),
            'mot_cle_client_url': clientKpi.urlsCount.toLocaleString('fr-FR'),
            'mot_cle_transac_client': intentStats.transac.posAll.toLocaleString('fr-FR'),
            'mot_cle_info_client': intentStats.info.posAll.toLocaleString('fr-FR'),
            'mot_cle_transac_client_top_10': intentStats.transac.top10.toLocaleString('fr-FR'),
            'mot_cle_info_client_top_10': intentStats.info.top10.toLocaleString('fr-FR'),
            'mot_cle_transac_pct': Math.round(transacPctDec * 100) + "%",
            'mot_cle_info_pct': Math.round(infoPctDec * 100) + "%"
        };

        slides.forEach(function(slide) {
            var shapes = slide.getShapes();

            // Étape A : Remplacement des textes classiques
            shapes.forEach(function(shape) {
                var shapeText = shape.getText().asString().trim();
                var titleRaw = shape.getTitle() || "";
                var descRaw = shape.getDescription() || "";

                // On vérifie le texte, le titre de la forme ou la description
                var targetKey = mapping[shapeText] ? shapeText : (mapping[titleRaw] ? titleRaw : (mapping[descRaw] ? descRaw : null));

                if (targetKey) {
                    shape.getText().setText(mapping[targetKey].toString());
                }
            });

            // Étape B : Traitement des jauges dynamiques (Conservation de l'arrondi)
            var shapesForGauges = slide.getShapes();
            shapesForGauges.forEach(function(shape) {
                var shapeText = shape.getText().asString().trim();
                var titleRaw = shape.getTitle() || "";
                var descRaw = shape.getDescription() || "";
                
                var isTransacGauge = (titleRaw === "jauge_transac_top10" || descRaw === "jauge_transac_top10" || shapeText === "jauge_transac_top10");
                var isInfoGauge = (titleRaw === "jauge_info_top10" || descRaw === "jauge_info_top10" || shapeText === "jauge_info_top10");

                var targetGauge = isTransacGauge ? "transac" : (isInfoGauge ? "info" : null);

                if (targetGauge) {
                    var pct = (targetGauge === "transac") ? transacPctDec : infoPctDec;

                    var left = shape.getLeft();
                    var top = shape.getTop();
                    var width = shape.getWidth();

                    // 1. Dessiner la jauge verte en DUPLIQUANT la forme originale (conserve l'arrondi exact)
                    if (pct > 0) {
                        var fgShape = slide.insertShape(shape);
                        var fillWidth = Math.max(10, width * pct); // Minimum 10px pour que l'arrondi s'affiche correctement
                        fgShape.setWidth(fillWidth);
                        fgShape.setLeft(left);
                        fgShape.setTop(top); // Réalignement strict
                        fgShape.getFill().setSolidFill("#00b050"); // Vert
                        fgShape.getBorder().setTransparent();
                        fgShape.getText().clear();
                        fgShape.setTitle("");
                        fgShape.setDescription("");
                    }

                    // 2. Transformer la forme originale en fond gris
                    shape.getFill().setSolidFill("#f1f3f4"); // Gris clair
                    shape.getBorder().setTransparent();
                    shape.getText().clear();
                    shape.setTitle("");
                    shape.setDescription("");
                }
            });
        });

        Logger.log("=== FIN EXPORT SLIDE GLOBAL & INTENTIONS ===");
        return { success: true, url: presentation.getUrl() };

    } catch (e) {
        Logger.log("ERREUR CRITIQUE EXPORT GLOBAL : " + e.message);
        return { success: false, error: e.message };
    }
}
