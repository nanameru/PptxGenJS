/**
 * NAME: demo_ultra_text.mjs
 * AUTH: AI Assistant
 * DESC: Advanced JSON-to-PPTX conversion with rich table formatting
 * VER.: 1.0.0
 * BLD.: 20241221
 */

import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R } from "./enums.mjs";

/**
 * Generate advanced presentation from JSON data with sophisticated table formatting
 * @param {PptxGenJS} pptx - PptxGenJS instance
 */
export function genSlides_UltraText(pptx) {
    // JSON data from user request
    const jsonData = {
        "title": "コンサルティング提案 評価レポート",
        "author": "AI Slide Generator", 
        "company": "AI-Powered Presentation System",
        "slides": [
            {
                "slideIndex": 0,
                "topic": "コンサルティング提案 評価レポート",
                "elements": [
                    {
                        "options": { "h": 22, "w": 100, "x": 0, "y": 0 },
                        "shapeType": "rect",
                        "type": "shape",
                        "purpose": "background",
                        "style": { "fill": { "color": "#1A365D", "type": "solid" } }
                    },
                    {
                        "content": "コンサルティング提案 評価レポート",
                        "options": { "h": 8, "w": 90, "x": 5, "y": 4 },
                        "type": "text",
                        "style": {
                            "align": "left", "bold": true, "color": "#FFFFFF", "fontSize": 32
                        },
                        "textType": "title"
                    },
                    {
                        "content": "A社向けDX戦略提案：実現性評価とパートナー選定",
                        "options": { "h": 5, "w": 90, "x": 5, "y": 13 },
                        "type": "text",
                        "style": {
                            "align": "left", "bold": false, "color": "#E0E0E0", "fontSize": 18
                        },
                        "textType": "subtitle"
                    },
                    {
                        "data": [
                            [
                                {
                                    "text": "評価項目",
                                    "options": {
                                        "align": "center", "bold": true, "color": "FFFFFF", "fill": "4A5568"
                                    }
                                },
                                {
                                    "text": "コンサルティングA社",
                                    "options": {
                                        "align": "center", "bold": true, "color": "FFFFFF", "fill": "4A5568"
                                    }
                                },
                                {
                                    "text": "コンサルティングB社",
                                    "options": {
                                        "align": "center", "bold": true, "color": "FFFFFF", "fill": "2B4A87"
                                    }
                                },
                                {
                                    "text": "コンサルティングC社",
                                    "options": {
                                        "align": "center", "bold": true, "color": "FFFFFF", "fill": "4A5568"
                                    }
                                }
                            ],
                            [
                                {
                                    "text": "提案コスト (百万円)",
                                    "options": { "bold": true, "fill": "F7FAFC" }
                                },
                                { "text": "2,500", "options": { "align": "right" } },
                                {
                                    "text": "2,200",
                                    "options": { "align": "right", "bold": true, "fill": "E6F3FF" }
                                },
                                { "text": "3,000", "options": { "align": "right" } }
                            ],
                            [
                                {
                                    "text": "期待ROI (%)",
                                    "options": { "bold": true, "fill": "F7FAFC" }
                                },
                                { "text": "150", "options": { "align": "right" } },
                                {
                                    "text": "185",
                                    "options": { "align": "right", "bold": true, "fill": "E6F3FF" }
                                },
                                { "text": "140", "options": { "align": "right" } }
                            ],
                            [
                                {
                                    "text": "専門性評価 (5段階)",
                                    "options": { "bold": true, "fill": "F7FAFC" }
                                },
                                { "text": "4.5", "options": { "align": "right" } },
                                {
                                    "text": "4.8",
                                    "options": { "align": "right", "bold": true, "fill": "E6F3FF" }
                                },
                                { "text": "4.2", "options": { "align": "right" } }
                            ]
                        ],
                        "options": { "h": 38, "w": 90, "x": 5, "y": 32 },
                        "type": "table",
                        "colW": [30, 20, 20, 20],
                        "title": "コンサルティング3社 比較評価"
                    }
                ],
                "background": "#F8F9FA"
            }
        ]
    };

    // Set presentation metadata
    pptx.title = jsonData.title;
    pptx.author = jsonData.author;
    pptx.company = jsonData.company;
    pptx.layout = 'LAYOUT_16x9';

    pptx.addSection({ title: "Ultra Rich Tables Demo" });

    // Generate slides from JSON data
    jsonData.slides.forEach((slideData, slideIndex) => {
        const slide = pptx.addSlide({ sectionTitle: "Ultra Rich Tables Demo" });
        
        // Set slide background
        if (slideData.background) {
            slide.background = { color: slideData.background.replace('#', '') };
        }

        slideData.elements.forEach(element => {
            // Convert JSON percentage coordinates to PptxGenJS inches (approximate)
            const convertCoord = (value, isSize = false) => {
                if (typeof value === 'number') {
                    return isSize ? value / 10 : value / 10; // Convert percentage to rough inch equivalent
                }
                return value;
            };

            const opts = {
                x: convertCoord(element.options?.x || 0),
                y: convertCoord(element.options?.y || 0),
                w: convertCoord(element.options?.w || 10, true),
                h: convertCoord(element.options?.h || 5, true)
            };

            switch (element.type) {
                case 'shape':
                    const shapeOpts = {
                        ...opts,
                        fill: element.style?.fill ? { color: element.style.fill.color.replace('#', '') } : { color: 'CCCCCC' }
                    };
                    
                    // Map shapeType from JSON to PptxGenJS enum
                    let shapeType = 'rect'; // default
                    if (element.shapeType === 'roundedRectangle') shapeType = 'roundRect';
                    else if (element.shapeType) shapeType = element.shapeType;
                    
                    slide.addShape(pptx.shapes[shapeType] || pptx.shapes.rect, shapeOpts);
                    break;

                case 'text':
                    const textOpts = {
                        ...opts,
                        fontSize: element.style?.fontSize || 16,
                        color: element.style?.color ? element.style.color.replace('#', '') : '000000',
                        bold: element.style?.bold || false,
                        align: element.style?.align || 'left'
                    };
                    slide.addText(element.content, textOpts);
                    break;

                case 'table':
                    // **SOPHISTICATED TABLE PROCESSING**
                    const tableData = processTableData(element.data);
                    const tableOpts = {
                        ...opts,
                        colW: element.colW ? element.colW.map(w => w / 10) : undefined, // Convert percentage to inches
                        fontSize: 12,
                        border: { pt: 1, color: 'CCCCCC' },
                        align: 'center',
                        valign: 'middle'
                    };
                    
                    slide.addTable(tableData, tableOpts);
                    break;
            }
        });
    });

    // Add additional demo slides showing table capabilities
    addAdvancedTableDemos(pptx);
}

/**
 * Process complex table data from JSON to PptxGenJS format
 * This function handles cell-level styling, colors, alignment, etc.
 */
function processTableData(jsonTableData) {
    if (!Array.isArray(jsonTableData)) return [];

    return jsonTableData.map(row => {
        if (!Array.isArray(row)) return [];
        
        return row.map(cell => {
            // Handle both simple strings and complex cell objects
            if (typeof cell === 'string') {
                return { text: cell };
            }
            
            if (typeof cell === 'object' && cell.text !== undefined) {
                const cellObj = {
                    text: cell.text,
                    options: {}
                };

                // Process cell options if they exist
                if (cell.options) {
                    const opts = cell.options;
                    
                    // Text alignment
                    if (opts.align) cellObj.options.align = opts.align;
                    if (opts.valign) cellObj.options.valign = opts.valign;
                    
                    // Text styling
                    if (opts.bold) cellObj.options.bold = opts.bold;
                    if (opts.italic) cellObj.options.italic = opts.italic;
                    if (opts.color) cellObj.options.color = opts.color.replace('#', '');
                    if (opts.fontSize) cellObj.options.fontSize = opts.fontSize;
                    
                    // Cell background fill
                    if (opts.fill) {
                        cellObj.options.fill = { color: opts.fill.replace('#', '') };
                    }
                    
                    // Cell spanning
                    if (opts.colspan) cellObj.options.colspan = opts.colspan;
                    if (opts.rowspan) cellObj.options.rowspan = opts.rowspan;
                    
                    // Cell margins/padding
                    if (opts.margin) cellObj.options.margin = opts.margin;
                    
                    // Cell borders (if specified)
                    if (opts.border) {
                        if (Array.isArray(opts.border)) {
                            cellObj.options.border = opts.border.map(border => ({
                                type: border.type || 'solid',
                                color: border.color || '000000',
                                pt: border.pt || 1
                            }));
                        } else {
                            cellObj.options.border = {
                                type: opts.border.type || 'solid',
                                color: opts.border.color || '000000',
                                pt: opts.border.pt || 1
                            };
                        }
                    }
                }

                return cellObj;
            }
            
            // Fallback for unexpected data
            return { text: String(cell) };
        });
    });
}

/**
 * Add additional demonstration slides showing advanced table features
 */
function addAdvancedTableDemos(pptx) {
    // SLIDE: Advanced Cell Styling Demo
    {
        const slide = pptx.addSlide({ sectionTitle: "Ultra Rich Tables Demo" });
        slide.addText("Advanced Cell Styling & Borders", {
            x: 0.5, y: 0.5, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1A365D'
        });

        const advancedTableData = [
            [
                { text: "カテゴリ", options: { bold: true, fill: "2B4A87", color: "FFFFFF", align: "center" } },
                { text: "現状", options: { bold: true, fill: "4A5568", color: "FFFFFF", align: "center" } },
                { text: "目標", options: { bold: true, fill: "0D47A1", color: "FFFFFF", align: "center" } },
                { text: "改善率", options: { bold: true, fill: "1565C0", color: "FFFFFF", align: "center" } }
            ],
            [
                { text: "効率性", options: { bold: true, fill: "F7FAFC" } },
                { text: "65%", options: { align: "center" } },
                { text: "85%", options: { align: "center", bold: true, color: "2E7D32" } },
                { text: "+31%", options: { align: "center", bold: true, fill: "C8E6C9", color: "1B5E20" } }
            ],
            [
                { text: "満足度", options: { bold: true, fill: "F7FAFC" } },
                { text: "7.2", options: { align: "center" } },
                { text: "9.1", options: { align: "center", bold: true, color: "2E7D32" } },
                { text: "+26%", options: { align: "center", bold: true, fill: "C8E6C9", color: "1B5E20" } }
            ],
            [
                { text: "コスト削減", options: { bold: true, fill: "F7FAFC" } },
                { text: "¥120M", options: { align: "center" } },
                { text: "¥180M", options: { align: "center", bold: true, color: "2E7D32" } },
                { text: "+50%", options: { align: "center", bold: true, fill: "C8E6C9", color: "1B5E20" } }
            ]
        ];

        slide.addTable(advancedTableData, {
            x: 0.5, y: 1.5, w: 12, h: 4,
            colW: [3, 2.5, 2.5, 4],
            fontSize: 14,
            border: { pt: 1, color: 'D1D5DB' },
            align: 'center',
            valign: 'middle'
        });
    }

    // SLIDE: Complex Multi-level Headers Demo  
    {
        const slide = pptx.addSlide({ sectionTitle: "Ultra Rich Tables Demo" });
        slide.addText("Multi-level Headers & Merged Cells", {
            x: 0.5, y: 0.5, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1A365D'
        });

        const multiLevelData = [
            [
                { text: "製品カテゴリ", options: { rowspan: 2, bold: true, fill: "1565C0", color: "FFFFFF", align: "center", valign: "middle" } },
                { text: "2023年実績", options: { colspan: 2, bold: true, fill: "2196F3", color: "FFFFFF", align: "center" } },
                { text: "2024年予測", options: { colspan: 2, bold: true, fill: "42A5F5", color: "FFFFFF", align: "center" } }
            ],
            [
                // rowspan cell from previous row
                { text: "売上", options: { bold: true, fill: "E3F2FD", align: "center" } },
                { text: "利益", options: { bold: true, fill: "E3F2FD", align: "center" } },
                { text: "売上", options: { bold: true, fill: "E8F5E8", align: "center" } },
                { text: "利益", options: { bold: true, fill: "E8F5E8", align: "center" } }
            ],
            [
                { text: "ソフトウェア", options: { bold: true, fill: "F5F5F5" } },
                { text: "¥450M", options: { align: "right" } },
                { text: "¥85M", options: { align: "right" } },
                { text: "¥520M", options: { align: "right", bold: true, color: "2E7D32" } },
                { text: "¥105M", options: { align: "right", bold: true, color: "2E7D32" } }
            ],
            [
                { text: "ハードウェア", options: { bold: true, fill: "F5F5F5" } },
                { text: "¥320M", options: { align: "right" } },
                { text: "¥48M", options: { align: "right" } },
                { text: "¥380M", options: { align: "right", bold: true, color: "2E7D32" } },
                { text: "¥65M", options: { align: "right", bold: true, color: "2E7D32" } }
            ],
            [
                { text: "サービス", options: { bold: true, fill: "F5F5F5" } },
                { text: "¥280M", options: { align: "right" } },
                { text: "¥42M", options: { align: "right" } },
                { text: "¥350M", options: { align: "right", bold: true, color: "2E7D32" } },
                { text: "¥58M", options: { align: "right", bold: true, color: "2E7D32" } }
            ]
        ];

        slide.addTable(multiLevelData, {
            x: 0.5, y: 1.5, w: 12, h: 5,
            colW: [3, 2, 2, 2.5, 2.5],
            fontSize: 13,
            border: { pt: 1, color: 'BDBDBD' },
            align: 'center',
            valign: 'middle'
        });
    }

    // SLIDE: Financial Dashboard Style Table
    {
        const slide = pptx.addSlide({ sectionTitle: "Ultra Rich Tables Demo" });
        slide.addText("Financial Dashboard Style", {
            x: 0.5, y: 0.5, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1A365D'
        });

        const financialData = [
            [
                { text: "指標", options: { bold: true, fill: "263238", color: "FFFFFF", align: "center" } },
                { text: "Q1", options: { bold: true, fill: "37474F", color: "FFFFFF", align: "center" } },
                { text: "Q2", options: { bold: true, fill: "455A64", color: "FFFFFF", align: "center" } },
                { text: "Q3", options: { bold: true, fill: "546E7A", color: "FFFFFF", align: "center" } },
                { text: "Q4", options: { bold: true, fill: "607D8B", color: "FFFFFF", align: "center" } },
                { text: "YoY成長", options: { bold: true, fill: "1B5E20", color: "FFFFFF", align: "center" } }
            ],
            [
                { text: "売上高", options: { bold: true, fill: "ECEFF1" } },
                { text: "¥2.1B", options: { align: "right", fontSize: 12 } },
                { text: "¥2.3B", options: { align: "right", fontSize: 12 } },
                { text: "¥2.7B", options: { align: "right", fontSize: 12 } },
                { text: "¥2.9B", options: { align: "right", fontSize: 12, bold: true } },
                { text: "+18%", options: { align: "center", bold: true, color: "2E7D32", fill: "E8F5E8" } }
            ],
            [
                { text: "営業利益", options: { bold: true, fill: "ECEFF1" } },
                { text: "¥420M", options: { align: "right", fontSize: 12 } },
                { text: "¥485M", options: { align: "right", fontSize: 12 } },
                { text: "¥540M", options: { align: "right", fontSize: 12 } },
                { text: "¥580M", options: { align: "right", fontSize: 12, bold: true } },
                { text: "+25%", options: { align: "center", bold: true, color: "2E7D32", fill: "E8F5E8" } }
            ],
            [
                { text: "EBITDA", options: { bold: true, fill: "ECEFF1" } },
                { text: "¥520M", options: { align: "right", fontSize: 12 } },
                { text: "¥595M", options: { align: "right", fontSize: 12 } },
                { text: "¥650M", options: { align: "right", fontSize: 12 } },
                { text: "¥720M", options: { align: "right", fontSize: 12, bold: true } },
                { text: "+31%", options: { align: "center", bold: true, color: "2E7D32", fill: "E8F5E8" } }
            ],
            [
                { text: "利益率", options: { bold: true, fill: "ECEFF1" } },
                { text: "20.0%", options: { align: "right", fontSize: 12 } },
                { text: "21.1%", options: { align: "right", fontSize: 12 } },
                { text: "20.0%", options: { align: "right", fontSize: 12 } },
                { text: "20.0%", options: { align: "right", fontSize: 12, bold: true } },
                { text: "+1.1%", options: { align: "center", bold: true, color: "2E7D32", fill: "E8F5E8" } }
            ]
        ];

        slide.addTable(financialData, {
            x: 0.5, y: 1.5, w: 12, h: 4,
            colW: [2.5, 2, 2, 2, 2, 1.5],
            fontSize: 11,
            border: { pt: 0.5, color: 'CFD8DC' },
            align: 'center',
            valign: 'middle'
        });

        // Add summary callout box
        slide.addShape(pptx.shapes.roundRect, {
            x: 1, y: 6, w: 10, h: 1.2,
            fill: { color: 'FFF3E0' },
            line: { color: 'FF9800', width: 1 }
        });

        slide.addText("💡 総評：全四半期にわたって持続的な成長を実現。特にQ4の業績が目標を20%上回る好結果。", {
            x: 1.2, y: 6.2, w: 9.6, h: 0.8,
            fontSize: 14, bold: false, color: 'E65100', align: 'left', valign: 'middle'
        });
    }
} 