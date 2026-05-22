document.addEventListener("DOMContentLoaded", () => {
    // DOM Elements
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const dashboard = document.getElementById("dashboard");
    const currentFilename = document.getElementById("current-filename");
    const btnReupload = document.getElementById("btn-reupload");
    const btnCopy = document.getElementById("btn-copy");
    const btnDownload = document.getElementById("btn-download");
    const yamlOutput = document.getElementById("yaml-output");
    const toast = document.getElementById("toast");

    // Overview Elements
    const metaSlides = document.getElementById("meta-slides");
    const metaRatio = document.getElementById("meta-ratio");
    const metaPx = document.getElementById("meta-px");
    const metaInches = document.getElementById("meta-inches");
    const metaCreator = document.getElementById("meta-creator");
    const metaCreated = document.getElementById("meta-created");
    const metaModified = document.getElementById("meta-modified");

    // Themes Elements
    const mainColorsGrid = document.getElementById("main-colors-grid");
    const accentColorsGrid = document.getElementById("accent-colors-grid");
    const fontMajorLatin = document.getElementById("font-major-latin");
    const fontMajorEa = document.getElementById("font-major-ea");
    const fontMinorLatin = document.getElementById("font-minor-latin");
    const fontMinorEa = document.getElementById("font-minor-ea");
    const inferredStylesList = document.getElementById("inferred-styles-list");

    // State Variables
    let currentYamlData = "";
    let currentYamlObj = null;

    // Drag & Drop Event Listeners
    dropZone.addEventListener("click", () => fileInput.click());
    
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });

    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("dragover");
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    btnReupload.addEventListener("click", () => {
        dashboard.classList.add("hidden");
        dropZone.classList.remove("hidden");
        fileInput.value = "";
    });

    // Copy to Clipboard
    btnCopy.addEventListener("click", () => {
        if (!currentYamlData) return;
        navigator.clipboard.writeText(currentYamlData).then(() => {
            showToast("複製成功！");
        }).catch(err => {
            console.error("無法複製到剪貼簿", err);
            showToast("複製失敗，請手動選取複製");
        });
    });

    // Download YAML file
    btnDownload.addEventListener("click", () => {
        if (!currentYamlData) return;
        const blob = new Blob([currentYamlData], { type: "text/yaml;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        const filename = (currentFilename.textContent.replace(/\.[^/.]+$/, "")) + "_style.yaml";
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast("下載完成！");
    });

    // Main File Handler
    function handleFile(file) {
        if (!file.name.endsWith(".pptx")) {
            alert("請選擇標準的 PPTX 簡報檔案！");
            return;
        }

        currentFilename.textContent = file.name;
        
        // Show loading state in dropZone
        dropZone.querySelector("h2").textContent = "正在解析簡報中...";
        dropZone.querySelector("p").textContent = "請稍候，這完全在本地瀏覽器中處理";

        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const arrayBuffer = e.target.result;
                const zip = await JSZip.loadAsync(arrayBuffer);
                await parsePPTX(zip);
                
                // Show dashboard and hide upload
                dropZone.classList.add("hidden");
                dashboard.classList.remove("hidden");
            } catch (err) {
                console.error("解析錯誤:", err);
                alert("無法解析此 PPTX 檔案，請確認其格式是否正確。詳細錯誤: " + err.message);
                // Reset upload button
                dropZone.querySelector("h2").textContent = "將 PPTX 簡報拖曳至此處";
                dropZone.querySelector("p").textContent = "或者點擊此處選擇您電腦中的檔案";
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // Toast Notification helper
    function showToast(message) {
        toast.textContent = message;
        toast.classList.remove("hidden");
        setTimeout(() => {
            toast.classList.add("hidden");
        }, 2000);
    }

    // Helper functions for XML Parsing (Namespace Agnostic)
    function getElementByLocalName(doc, parent, localName) {
        const parentNode = parent || doc;
        let el = parentNode.querySelector(localName);
        if (el) return el;
        
        // namespace selector fallback
        try {
            el = parentNode.querySelector(`*|${localName}`);
            if (el) return el;
        } catch(e) {}
        
        // tag lookup fallback
        const tags = parentNode.getElementsByTagName("*");
        for (let i = 0; i < tags.length; i++) {
            const tag = tags[i];
            const name = tag.localName || tag.tagName.split(':').pop();
            if (name === localName) {
                return tag;
            }
        }
        return null;
    }

    function getElementsByLocalName(doc, parent, localName) {
        const parentNode = parent || doc;
        const result = [];
        const tags = parentNode.getElementsByTagName("*");
        for (let i = 0; i < tags.length; i++) {
            const tag = tags[i];
            const name = tag.localName || tag.tagName.split(':').pop();
            if (name === localName) {
                result.push(tag);
            }
        }
        return result;
    }

    // Mapped standard scheme color name conversions
    function mapSchemeColorName(name) {
        const maps = {
            "bg1": "lt1",
            "bg2": "lt2",
            "tx1": "dk1",
            "tx2": "dk2"
        };
        return maps[name] || name;
    }

    // Adjust luminance (tint/shade) based on lumMod and lumOff attributes
    function adjustLuminance(hex, mod, off) {
        if (!hex || hex.length !== 7 || !hex.startsWith("#")) return hex;
        let r = parseInt(hex.slice(1, 3), 16);
        let g = parseInt(hex.slice(3, 5), 16);
        let b = parseInt(hex.slice(5, 7), 16);
        
        // Apply modification: val * mod + off
        r = Math.min(255, Math.max(0, Math.round(r * mod + off * 255)));
        g = Math.min(255, Math.max(0, Math.round(g * mod + off * 255)));
        b = Math.min(255, Math.max(0, Math.round(b * mod + off * 255)));
        
        const toHex = (n) => n.toString(16).padStart(2, '0');
        return `#${toHex(r)}${toHex(g)}${toHex(b)}`.toUpperCase();
    }

    // Resolve color nodes to actual Hex
    function resolveColor(themeColors, colorNode) {
        if (!colorNode) return null;
        
        let hex = null;
        const srgb = getElementByLocalName(null, colorNode, "srgbClr");
        if (srgb) {
            hex = "#" + srgb.getAttribute("val");
        } else {
            const sys = getElementByLocalName(null, colorNode, "sysClr");
            if (sys) {
                hex = "#" + (sys.getAttribute("lastClr") || "000000");
            }
        }

        const scheme = getElementByLocalName(null, colorNode, "schemeClr");
        if (scheme) {
            const val = scheme.getAttribute("val");
            const mappedName = mapSchemeColorName(val);
            if (themeColors && themeColors[mappedName]) {
                hex = themeColors[mappedName];
                
                // Read and apply lumMod / lumOff if present
                const lumMod = getElementByLocalName(null, scheme, "lumMod");
                const lumOff = getElementByLocalName(null, scheme, "lumOff");
                if (lumMod || lumOff) {
                    const modVal = lumMod ? parseInt(lumMod.getAttribute("val")) / 100000 : 1;
                    const offVal = lumOff ? parseInt(lumOff.getAttribute("val")) / 100000 : 0;
                    hex = adjustLuminance(hex, modVal, offVal);
                }
            }
        }
        
        return hex ? hex.toUpperCase() : null;
    }

    function getAspectRatio(w, h) {
        const r = w / h;
        if (Math.abs(r - 1.7777777777777777) < 0.02) return "16:9";
        if (Math.abs(r - 1.3333333333333333) < 0.02) return "4:3";
        if (Math.abs(r - 1.6) < 0.02) return "16:10";
        
        // Find greatest common divisor
        const gcd = (a, b) => b ? gcd(b, a % b) : a;
        const wInt = Math.round(w * 100);
        const hInt = Math.round(h * 100);
        const d = gcd(wInt, hInt);
        return `${wInt / d}:${hInt / d}`;
    }

    function containsChinese(str) {
        return /[\u4e00-\u9fa5]/.test(str);
    }

    // PPTX XML Parser Logic
    async function parsePPTX(zip) {
        const parser = new DOMParser();

        // 1. Parse Metadata (docProps/core.xml & docProps/app.xml)
        let creator = "Unknown";
        let createdAt = "Unknown";
        let modifiedAt = "Unknown";
        
        const coreFile = zip.file("docProps/core.xml");
        if (coreFile) {
            const coreText = await coreFile.async("text");
            const coreDoc = parser.parseFromString(coreText, "application/xml");
            
            const creatorEl = getElementByLocalName(coreDoc, null, "creator");
            const createdEl = getElementByLocalName(coreDoc, null, "created");
            const modifiedEl = getElementByLocalName(coreDoc, null, "modified");
            
            if (creatorEl) creator = creatorEl.textContent;
            if (createdEl) createdAt = new Date(createdEl.textContent).toLocaleString('zh-TW');
            if (modifiedAt) modifiedAt = new Date(modifiedEl.textContent).toLocaleString('zh-TW');
        }

        // Count Slides
        const slideFiles = zip.file(/ppt\/slides\/slide[0-9]+\.xml/);
        const slideCount = slideFiles.length;

        // 2. Parse Presentation Properties (ppt/presentation.xml)
        let dimensions = {
            width_px: 1920,
            height_px: 1080,
            width_inches: 13.33,
            height_inches: 7.5,
            aspect_ratio: "16:9"
        };
        
        const presFile = zip.file("ppt/presentation.xml");
        if (presFile) {
            const presText = await presFile.async("text");
            const presDoc = parser.parseFromString(presText, "application/xml");
            const sldSz = getElementByLocalName(presDoc, null, "sldSz");
            if (sldSz) {
                const cx = parseInt(sldSz.getAttribute("cx") || "0");
                const cy = parseInt(sldSz.getAttribute("cy") || "0");
                if (cx > 0 && cy > 0) {
                    dimensions.width_inches = parseFloat((cx / 914400).toFixed(2));
                    dimensions.height_inches = parseFloat((cy / 914400).toFixed(2));
                    dimensions.width_px = Math.round(cx / 9525);
                    dimensions.height_px = Math.round(cy / 9525);
                    dimensions.aspect_ratio = getAspectRatio(dimensions.width_inches, dimensions.height_inches);
                }
            }
        }

        // 3. Parse Theme (ppt/theme/theme1.xml)
        let themeName = "Office Theme";
        let themeColors = {
            dk1: "#000000", lt1: "#FFFFFF",
            dk2: "#1F497D", lt2: "#EEECE1",
            accent1: "#4F81BD", accent2: "#C0504D",
            accent3: "#9BBB59", accent4: "#8064A2",
            accent5: "#4BACC6", accent6: "#F79646",
            hlink: "#0000FF", folHlink: "#800080"
        };
        let themeFonts = {
            major: { latin: "Calibri Light", ea: "Microsoft JhengHei" },
            minor: { latin: "Calibri", ea: "Microsoft JhengHei" }
        };

        const themeFile = zip.file("ppt/theme/theme1.xml");
        if (themeFile) {
            const themeText = await themeFile.async("text");
            const themeDoc = parser.parseFromString(themeText, "application/xml");
            
            // Name
            const themeEl = getElementByLocalName(themeDoc, null, "theme");
            if (themeEl && themeEl.getAttribute("name")) {
                themeName = themeEl.getAttribute("name");
            }

            // Colors
            const clrScheme = getElementByLocalName(themeDoc, null, "clrScheme");
            if (clrScheme) {
                const colorKeys = Object.keys(themeColors);
                for (const key of colorKeys) {
                    const node = getElementByLocalName(themeDoc, clrScheme, key);
                    if (node) {
                        const resolved = resolveColor(null, node);
                        if (resolved) {
                            themeColors[key] = resolved;
                        }
                    }
                }
            }

            // Fonts
            const fontScheme = getElementByLocalName(themeDoc, null, "fontScheme");
            if (fontScheme) {
                const majorNode = getElementByLocalName(themeDoc, fontScheme, "majorFont");
                const minorNode = getElementByLocalName(themeDoc, fontScheme, "minorFont");
                
                const parseFonts = (node) => {
                    if (!node) return null;
                    const latin = getElementByLocalName(themeDoc, node, "latin");
                    const ea = getElementByLocalName(themeDoc, node, "ea");
                    return {
                        latin: latin ? latin.getAttribute("typeface") : "",
                        ea: ea ? ea.getAttribute("typeface") : ""
                    };
                };

                const parsedMajor = parseFonts(majorNode);
                if (parsedMajor) {
                    themeFonts.major.latin = parsedMajor.latin || themeFonts.major.latin;
                    themeFonts.major.ea = parsedMajor.ea || themeFonts.major.ea;
                }
                
                const parsedMinor = parseFonts(minorNode);
                if (parsedMinor) {
                    themeFonts.minor.latin = parsedMinor.latin || themeFonts.minor.latin;
                    themeFonts.minor.ea = parsedMinor.ea || themeFonts.minor.ea;
                }
            }
        }

        // 4. Slide Background & Text Style Extraction (Frequency Analysis)
        let backgroundOccurrence = {};
        let shapeFills = new Set();
        let shapeBorders = new Set();
        let imageCount = 0;

        // Typography styles frequency storage
        // Structure: { categoryName: { key: { styleProps, count } } }
        let typographyStats = {
            title: {},
            subtitle: {},
            body: {},
            caption: {}
        };

        const slidesData = [];

        for (const file of slideFiles) {
            const slideText = await file.async("text");
            const slideDoc = parser.parseFromString(slideText, "application/xml");

            // Get slide index
            const slideNameMatch = file.name.match(/slide([0-9]+)\.xml/);
            const slideIndex = slideNameMatch ? parseInt(slideNameMatch[1]) : 0;

            // Load relationship file for resolving image target paths
            let relsMap = {};
            const relsFileName = `ppt/slides/_rels/slide${slideIndex}.xml.rels`;
            const relsFile = zip.file(relsFileName);
            if (relsFile) {
                const relsText = await relsFile.async("text");
                const relsDoc = parser.parseFromString(relsText, "application/xml");
                const relationships = getElementsByLocalName(relsDoc, null, "Relationship");
                for (const rel of relationships) {
                    const id = rel.getAttribute("Id");
                    const target = rel.getAttribute("Target");
                    relsMap[id] = target;
                }
            }

            // Look for Background
            const bg = getElementByLocalName(slideDoc, null, "bg");
            if (bg) {
                const bgPr = getElementByLocalName(slideDoc, bg, "bgPr");
                if (bgPr) {
                    const color = resolveColor(themeColors, bgPr);
                    if (color) {
                        backgroundOccurrence[color] = (backgroundOccurrence[color] || 0) + 1;
                    } else if (getElementByLocalName(slideDoc, bgPr, "blipFill")) {
                        backgroundOccurrence["Image / Pattern Background"] = (backgroundOccurrence["Image / Pattern Background"] || 0) + 1;
                    }
                }
            } else {
                // Default to white if not specified
                backgroundOccurrence["#FFFFFF"] = (backgroundOccurrence["#FFFFFF"] || 0) + 1;
            }

            // Look for Images
            const blips = getElementsByLocalName(slideDoc, null, "blip");
            imageCount += blips.length;

            // Look for Shapes
            const shapes = getElementsByLocalName(slideDoc, null, "sp");
            for (const sp of shapes) {
                const spPr = getElementByLocalName(slideDoc, sp, "spPr");
                if (spPr) {
                    const fill = resolveColor(themeColors, spPr);
                    if (fill) shapeFills.add(fill);

                    const ln = getElementByLocalName(slideDoc, spPr, "ln");
                    if (ln) {
                        const border = resolveColor(themeColors, ln);
                        if (border) shapeBorders.add(border);
                    }
                }

                // Analyze Text inside shape
                const txBody = getElementByLocalName(slideDoc, sp, "txBody");
                if (txBody) {
                    // Check if parent shape is a placeholder (Title, Subtitle, Body)
                    let inferredCategory = null;
                    const nvSpPr = getElementByLocalName(slideDoc, sp, "nvSpPr");
                    if (nvSpPr) {
                        const nvPr = getElementByLocalName(slideDoc, nvSpPr, "nvPr");
                        if (nvPr) {
                            const ph = getElementByLocalName(slideDoc, nvPr, "ph");
                            if (ph) {
                                const type = ph.getAttribute("type");
                                if (type === "title" || type === "ctrTitle") {
                                    inferredCategory = "title";
                                } else if (type === "subTitle") {
                                    inferredCategory = "subtitle";
                                } else if (type === "body") {
                                    inferredCategory = "body";
                                } else if (type === "dt" || type === "ftr" || type === "sldNum") {
                                    inferredCategory = "caption";
                                }
                            }
                        }
                    }

                    // Extract text runs inside txBody
                    const paragraphs = getElementsByLocalName(slideDoc, txBody, "p");
                    for (const pTag of paragraphs) {
                        const pPr = getElementByLocalName(slideDoc, pTag, "pPr");
                        const alignVal = pPr ? pPr.getAttribute("algn") : null;
                        
                        // Map alignVal to standard CSS/YAML string
                        let alignment = "left";
                        if (alignVal === "ctr") alignment = "center";
                        if (alignVal === "r") alignment = "right";
                        if (alignVal === "just") alignment = "justify";

                        const runs = getElementsByLocalName(slideDoc, pTag, "r");
                        for (const r of runs) {
                            const t = getElementByLocalName(slideDoc, r, "t");
                            if (!t || !t.textContent.trim()) continue;

                            const rPr = getElementByLocalName(slideDoc, r, "rPr");
                            let fontSize = 18; // default fallback
                            let bold = false;
                            let color = "#000000"; // default fallback
                            let fontFamily = themeFonts.minor.latin;

                            const textVal = t.textContent;
                            const isChinese = containsChinese(textVal);

                            if (rPr) {
                                // Font Size: sz is in hundredths of a point
                                if (rPr.getAttribute("sz")) {
                                    fontSize = Math.round(parseInt(rPr.getAttribute("sz")) / 100);
                                }
                                // Bold
                                bold = rPr.getAttribute("b") === "1";
                                
                                // Color
                                const resolvedC = resolveColor(themeColors, rPr);
                                if (resolvedC) color = resolvedC;

                                // Font Family
                                if (isChinese) {
                                    const ea = getElementByLocalName(slideDoc, rPr, "ea");
                                    fontFamily = ea ? ea.getAttribute("typeface") : themeFonts.minor.ea;
                                } else {
                                    const latin = getElementByLocalName(slideDoc, rPr, "latin");
                                    fontFamily = latin ? latin.getAttribute("typeface") : themeFonts.minor.latin;
                                }
                            }

                            // If no placeholder inference, classify based on font size threshold
                            let category = inferredCategory;
                            if (!category) {
                                if (fontSize >= 32) category = "title";
                                else if (fontSize >= 20) category = "subtitle";
                                else if (fontSize >= 11) category = "body";
                                else category = "caption";
                            }

                            // Create key
                            const key = `${fontFamily}_${fontSize}_${color}_${bold}_${alignment}`;
                            
                            if (!typographyStats[category][key]) {
                                typographyStats[category][key] = {
                                    font_family: fontFamily,
                                    font_size_pt: fontSize,
                                    color: color,
                                    bold: bold,
                                    alignment: alignment,
                                    count: 0
                                };
                            }
                            typographyStats[category][key].count += textVal.length;
                        }
                    }
                }
            }

            // Extract pictures on this slide
            const pics = getElementsByLocalName(slideDoc, null, "pic");
            const slideImages = [];
            for (const pic of pics) {
                const cNvPr = getElementByLocalName(slideDoc, pic, "cNvPr");
                const blip = getElementByLocalName(slideDoc, pic, "blip");
                const xfrm = getElementByLocalName(slideDoc, pic, "xfrm");

                let imgId = "";
                let imgName = "";
                let altText = "";
                if (cNvPr) {
                    imgId = cNvPr.getAttribute("id") || "";
                    imgName = cNvPr.getAttribute("name") || "";
                    altText = cNvPr.getAttribute("descr") || "";
                }

                let rId = "";
                let sourcePath = "";
                if (blip) {
                    rId = blip.getAttribute("r:embed") || blip.getAttribute("embed") || "";
                    if (!rId) {
                        for (let i = 0; i < blip.attributes.length; i++) {
                            const attr = blip.attributes[i];
                            if (attr.nodeName.endsWith("embed")) {
                                rId = attr.nodeValue;
                                break;
                            }
                        }
                    }
                    if (rId && relsMap[rId]) {
                        sourcePath = relsMap[rId];
                        if (sourcePath.startsWith("../")) {
                            sourcePath = "ppt/" + sourcePath.substring(3);
                        }
                    }
                }

                let layout = {
                    x_inch: 0,
                    y_inch: 0,
                    width_inch: 0,
                    height_inch: 0,
                    rotation: 0
                };

                if (xfrm) {
                    const off = getElementByLocalName(slideDoc, xfrm, "off");
                    const ext = getElementByLocalName(slideDoc, xfrm, "ext");
                    const rotVal = xfrm.getAttribute("rot");

                    if (off) {
                        const x = parseInt(off.getAttribute("x") || "0");
                        const y = parseInt(off.getAttribute("y") || "0");
                        layout.x_inch = parseFloat((x / 914400).toFixed(2));
                        layout.y_inch = parseFloat((y / 914400).toFixed(2));
                    }
                    if (ext) {
                        const cx = parseInt(ext.getAttribute("cx") || "0");
                        const cy = parseInt(ext.getAttribute("cy") || "0");
                        layout.width_inch = parseFloat((cx / 914400).toFixed(2));
                        layout.height_inch = parseFloat((cy / 914400).toFixed(2));
                    }
                    if (rotVal) {
                        layout.rotation = Math.round(parseInt(rotVal) / 60000);
                    }
                }

                slideImages.push({
                    id: imgId,
                    name: imgName,
                    alt_text: altText,
                    source_path: sourcePath,
                    layout: layout
                });
            }

            slidesData.push({
                slide_number: slideIndex,
                images: slideImages
            });
        }

        // Sort slide data by slide number
        slidesData.sort((a, b) => a.slide_number - b.slide_number);

        // Deduce dominant style for each category
        const typographyFinal = {};
        for (const category of Object.keys(typographyStats)) {
            const styles = Object.values(typographyStats[category]);
            if (styles.length > 0) {
                // Sort by count descending
                styles.sort((a, b) => b.count - a.count);
                typographyFinal[category] = {
                    font_family: styles[0].font_family,
                    font_size_pt: styles[0].font_size_pt,
                    color: styles[0].color,
                    bold: styles[0].bold,
                    alignment: styles[0].alignment,
                    frequency_score: styles[0].count
                };
            } else {
                // Default placeholders
                if (category === "title") {
                    typographyFinal[category] = {
                        font_family: themeFonts.major.ea || themeFonts.major.latin,
                        font_size_pt: 40,
                        color: themeColors.dk2 || themeColors.dk1,
                        bold: true,
                        alignment: "left"
                    };
                } else if (category === "subtitle") {
                    typographyFinal[category] = {
                        font_family: themeFonts.minor.ea || themeFonts.minor.latin,
                        font_size_pt: 24,
                        color: themeColors.accent1,
                        bold: false,
                        alignment: "left"
                    };
                } else if (category === "body") {
                    typographyFinal[category] = {
                        font_family: themeFonts.minor.ea || themeFonts.minor.latin,
                        font_size_pt: 16,
                        color: themeColors.dk1,
                        bold: false,
                        alignment: "left"
                    };
                } else {
                    typographyFinal[category] = {
                        font_family: themeFonts.minor.ea || themeFonts.minor.latin,
                        font_size_pt: 11,
                        color: themeColors.lt2,
                        bold: false,
                        alignment: "left"
                    };
                }
            }
        }

        // Format Backgrounds Array
        const bgSummary = Object.entries(backgroundOccurrence).map(([bgType, count]) => {
            return {
                type: bgType.startsWith("#") ? "solid" : "custom_pattern",
                color: bgType.startsWith("#") ? bgType : null,
                occurrence_count: count
            };
        });

        // 5. Structure full output object
        const yamlObj = {
            metadata: {
                creator: creator,
                created_at: createdAt,
                modified_at: modifiedAt,
                slides_count: slideCount
            },
            layout: {
                dimensions: {
                    width_px: dimensions.width_px,
                    height_px: dimensions.height_px,
                    width_inches: dimensions.width_inches,
                    height_inches: dimensions.height_inches,
                    aspect_ratio: dimensions.aspect_ratio
                }
            },
            style_palette: {
                theme_name: themeName,
                colors: {
                    dark1: themeColors.dk1,
                    light1: themeColors.lt1,
                    dark2: themeColors.dk2,
                    light2: themeColors.lt2,
                    accent1: themeColors.accent1,
                    accent2: themeColors.accent2,
                    accent3: themeColors.accent3,
                    accent4: themeColors.accent4,
                    accent5: themeColors.accent5,
                    accent6: themeColors.accent6,
                    hyperlink: themeColors.hlink,
                    followed_hyperlink: themeColors.folHlink
                },
                fonts: {
                    major: {
                        latin: themeFonts.major.latin,
                        east_asian: themeFonts.major.ea
                    },
                    minor: {
                        latin: themeFonts.minor.latin,
                        east_asian: themeFonts.minor.ea
                    }
                }
            },
            typography: {
                slide_title: {
                    font_family: typographyFinal.title.font_family,
                    font_size_pt: typographyFinal.title.font_size_pt,
                    color: typographyFinal.title.color,
                    bold: typographyFinal.title.bold,
                    alignment: typographyFinal.title.alignment
                },
                subtitle: {
                    font_family: typographyFinal.subtitle.font_family,
                    font_size_pt: typographyFinal.subtitle.font_size_pt,
                    color: typographyFinal.subtitle.color,
                    bold: typographyFinal.subtitle.bold,
                    alignment: typographyFinal.subtitle.alignment
                },
                body_text: {
                    font_family: typographyFinal.body.font_family,
                    font_size_pt: typographyFinal.body.font_size_pt,
                    color: typographyFinal.body.color,
                    bold: typographyFinal.body.bold,
                    alignment: typographyFinal.body.alignment
                },
                caption: {
                    font_family: typographyFinal.caption.font_family,
                    font_size_pt: typographyFinal.caption.font_size_pt,
                    color: typographyFinal.caption.color,
                    bold: typographyFinal.caption.bold,
                    alignment: typographyFinal.caption.alignment
                }
            },
            backgrounds: bgSummary,
            visual_elements: {
                shapes: {
                    fill_colors: Array.from(shapeFills),
                    border_colors: Array.from(shapeBorders)
                },
                image_count: imageCount
            },
            slides: slidesData
        };

        currentYamlObj = yamlObj;
        currentYamlData = jsyaml.dump(yamlObj, { indent: 2, lineWidth: -1 });

        // Update UI
        updateUI(yamlObj, currentYamlData);
    }

    // Render data to UI
    function updateUI(yamlObj, yamlText) {
        // Overview
        metaSlides.textContent = yamlObj.metadata.slides_count;
        metaRatio.textContent = yamlObj.layout.dimensions.aspect_ratio;
        metaPx.textContent = `${yamlObj.layout.dimensions.width_px} × ${yamlObj.layout.dimensions.height_px}`;
        metaInches.textContent = `${yamlObj.layout.dimensions.width_inches}" × ${yamlObj.layout.dimensions.height_inches}"`;
        
        metaCreator.textContent = yamlObj.metadata.creator;
        metaCreated.textContent = yamlObj.metadata.created_at;
        metaModified.textContent = yamlObj.metadata.modified_at;

        // Render Color Grids
        mainColorsGrid.innerHTML = "";
        accentColorsGrid.innerHTML = "";

        const colors = yamlObj.style_palette.colors;
        const mainColorKeys = ["dark1", "light1", "dark2", "light2"];
        const accentColorKeys = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink", "followed_hyperlink"];

        const addColorChip = (container, label, hex) => {
            const chip = document.createElement("div");
            chip.className = "color-chip";
            chip.title = `點擊複製: ${hex}`;
            
            const preview = document.createElement("div");
            if (hex) {
                preview.className = "color-preview";
                preview.style.backgroundColor = hex;
            } else {
                preview.className = "color-preview-empty";
            }
            
            const meta = document.createElement("div");
            meta.className = "color-meta";
            
            const name = document.createElement("span");
            name.className = "color-name";
            name.textContent = label;
            
            const hexCode = document.createElement("span");
            hexCode.className = "color-hex";
            hexCode.textContent = hex || "N/A";
            
            meta.appendChild(name);
            meta.appendChild(hexCode);
            chip.appendChild(preview);
            chip.appendChild(meta);

            // Copy color on click
            chip.addEventListener("click", () => {
                if (hex) {
                    navigator.clipboard.writeText(hex).then(() => {
                        showToast(`已複製顏色 ${hex}`);
                    });
                }
            });
            
            container.appendChild(chip);
        };

        const colorLabelMap = {
            dark1: "主要深色 (D1)",
            light1: "主要淺色 (L1)",
            dark2: "次要深色 (D2)",
            light2: "次要淺色 (L2)",
            accent1: "強調色 1",
            accent2: "強調色 2",
            accent3: "強調色 3",
            accent4: "強調色 4",
            accent5: "強調色 5",
            accent6: "強調色 6",
            hyperlink: "超連結",
            followed_hyperlink: "已開啟連結"
        };

        mainColorKeys.forEach(key => addColorChip(mainColorsGrid, colorLabelMap[key], colors[key]));
        accentColorKeys.forEach(key => addColorChip(accentColorsGrid, colorLabelMap[key], colors[key]));

        // Fonts
        fontMajorLatin.textContent = yamlObj.style_palette.fonts.major.latin || "N/A";
        fontMajorEa.textContent = yamlObj.style_palette.fonts.major.east_asian || "N/A";
        fontMinorLatin.textContent = yamlObj.style_palette.fonts.minor.latin || "N/A";
        fontMinorEa.textContent = yamlObj.style_palette.fonts.minor.east_asian || "N/A";

        // Inferred Styles
        inferredStylesList.innerHTML = "";
        const roleLabelMap = {
            slide_title: "投影片標題",
            subtitle: "投影片副標題",
            body_text: "本文區塊",
            caption: "說明文字/頁尾"
        };

        Object.entries(yamlObj.typography).forEach(([role, style]) => {
            const item = document.createElement("div");
            item.className = "style-inferred-item";

            const meta = document.createElement("div");
            meta.className = "style-inferred-meta";

            const title = document.createElement("span");
            title.className = "style-inferred-role";
            title.innerHTML = `${roleLabelMap[role]} <span class="style-inferred-freq"> dominant </span>`;

            const specs = document.createElement("div");
            specs.className = "style-inferred-specs";
            
            // size
            const sizeTag = document.createElement("span");
            sizeTag.className = "style-spec-tag";
            sizeTag.textContent = `${style.font_size_pt}pt`;
            
            // font family
            const fontTag = document.createElement("span");
            fontTag.className = "style-spec-tag";
            fontTag.textContent = style.font_family;
            
            // bold
            const boldTag = document.createElement("span");
            boldTag.className = "style-spec-tag";
            boldTag.textContent = style.bold ? "粗體" : "一般";

            // align
            const alignTag = document.createElement("span");
            alignTag.className = "style-spec-tag";
            const alignIconMap = { left: "靠左", center: "居中", right: "靠右", justify: "分散" };
            alignTag.textContent = alignIconMap[style.alignment] || "靠左";

            specs.appendChild(sizeTag);
            specs.appendChild(fontTag);
            specs.appendChild(boldTag);
            specs.appendChild(alignTag);
            
            meta.appendChild(title);
            meta.appendChild(specs);

            const preview = document.createElement("div");
            preview.className = "style-inferred-preview";
            preview.style.fontFamily = style.font_family;
            preview.style.fontSize = `${Math.max(12, Math.min(24, style.font_size_pt))}px`;
            preview.style.fontWeight = style.bold ? "bold" : "normal";
            preview.style.color = style.color;
            preview.textContent = "樣式預覽 Sample";

            item.appendChild(meta);
            item.appendChild(preview);
            inferredStylesList.appendChild(item);
        });

        // Set YAML editor content with basic color highlighting wrapper
        yamlOutput.innerHTML = highlightYaml(yamlText);
    }

    // Micro HTML highlights generator for yaml output code box
    function highlightYaml(text) {
        return text
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            // Highlight keys
            .replace(/^([\s-]*)([\w_]+)(:)/gm, '$1<span class="yaml-key">$2</span>$3')
            // Highlight strings
            .replace(/(:\s+)(["'].*?["']|[^#\d\s[].*?)(?=\s*(?:#|$))/g, '$1<span class="yaml-string">$2</span>')
            // Highlight numbers
            .replace(/(:\s+)(\b\d+(?:\.\d+)?\b)/g, '$1<span class="yaml-number">$2</span>')
            // Highlight booleans
            .replace(/(:\s+)(true|false)/g, '$1<span class="yaml-boolean">$2</span>')
            // Highlight comments
            .replace(/(#.*)$/gm, '<span class="yaml-comment">$1</span>');
    }
});
