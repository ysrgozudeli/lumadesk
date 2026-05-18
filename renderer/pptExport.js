// PPT export — ported from peerluma's documentExport.ts (TS → JS).
// Uses pptxgenjs in renderer to build the .pptx, then ships the
// ArrayBuffer to main process via IPC for save dialog + write.
//
// Exposes: window.LumaPpt.exportToPptx({ title, content, theme, mermaidImages })
//          → ArrayBuffer
(function () {
  const SLIDE_W = 13.33;
  const MARGIN = 0.6;
  const CONTENT_W = SLIDE_W - MARGIN * 2;
  const INDENT_X = 0.8;
  const INDENT_W = SLIDE_W - INDENT_X * 2;

  const THEMES = {
    light: {
      bgColor: 'FFFFFF',
      titleBg: 'F8F9FA',
      titleColor: '1A1A1A',
      contentTitleColor: '1A1A1A',
      textColor: '333333',
      accentColor: '6B7280',
      tableBorderColor: 'E5E7EB',
      tableHeaderBg: 'F3F4F6',
      tableHeaderColor: '1A1A1A',
      quoteColor: '6B7280',
      subtitleColor: '6B7280',
    },
    dark: {
      bgColor: '1A1A2E',
      titleBg: '16213E',
      titleColor: 'FFFFFF',
      contentTitleColor: 'FFFFFF',
      textColor: 'E0E0E0',
      accentColor: '4A90D9',
      tableBorderColor: '2D2D44',
      tableHeaderBg: '16213E',
      tableHeaderColor: 'FFFFFF',
      quoteColor: '8899AA',
      subtitleColor: '8899AA',
    },
    branded: {
      bgColor: 'F8FAFC',
      titleBg: '3730A3',
      titleColor: 'FFFFFF',
      contentTitleColor: '312E81',
      textColor: '374151',
      accentColor: '4F46E5',
      tableBorderColor: 'C7D2FE',
      tableHeaderBg: 'EEF2FF',
      tableHeaderColor: '312E81',
      quoteColor: '6366F1',
      subtitleColor: '6366F1',
    },
  };

  // PptxGenJS is exposed globally by the UMD bundle loaded in index.html.
  let marked = null;

  function getPptxGenJS() {
    if (typeof window.PptxGenJS !== 'function') {
      throw new Error('PptxGenJS bundle not loaded — check index.html script order');
    }
    return window.PptxGenJS;
  }

  async function loadMarked() {
    if (!marked) {
      const mod = await import('../node_modules/marked/lib/marked.esm.js');
      marked = mod.marked;
    }
    return marked;
  }

  function parseInlineForPptx(text, baseOptions = {}) {
    const segments = [];
    let remaining = text || '';

    remaining = remaining.replace(/\*\*(.+?)\*\*/g, '{{BOLD:$1}}');
    remaining = remaining.replace(/\*(.+?)\*/g, '{{ITALIC:$1}}');
    remaining = remaining.replace(/`(.+?)`/g, '{{CODE:$1}}');

    const tokenRegex = /\{\{(BOLD|ITALIC|CODE):([^}]+)\}\}/g;
    let lastIndex = 0;
    let m;

    while ((m = tokenRegex.exec(remaining)) !== null) {
      if (m.index > lastIndex) {
        segments.push({ text: remaining.slice(lastIndex, m.index), options: { ...baseOptions } });
      }
      const [, type, content] = m;
      if (type === 'BOLD') segments.push({ text: content, options: { ...baseOptions, bold: true } });
      else if (type === 'ITALIC') segments.push({ text: content, options: { ...baseOptions, italic: true } });
      else if (type === 'CODE') segments.push({ text: content, options: { ...baseOptions, fontFace: 'Courier New', fontSize: 10 } });
      lastIndex = m.index + m[0].length;
    }

    if (lastIndex < remaining.length) {
      segments.push({ text: remaining.slice(lastIndex), options: { ...baseOptions } });
    }
    if (segments.length === 0) {
      segments.push({ text: text || '', options: { ...baseOptions } });
    }
    return segments;
  }

  function addContentToSlide(slide, content, theme, startY, mermaidImages, markedLib) {
    const tokens = markedLib.lexer(content);
    let yPos = startY;

    for (let tIdx = 0; tIdx < tokens.length; tIdx++) {
      const token = tokens[tIdx];
      if (yPos > 6.8) break;

      switch (token.type) {
        case 'heading': {
          const fontSize = token.depth === 3 ? 18 : 16;
          const headingSegs = parseInlineForPptx(token.text || '', {
            fontSize, bold: true, color: theme.textColor,
          });
          slide.addText(headingSegs, {
            x: MARGIN, y: yPos, w: CONTENT_W, h: 0.4,
            fontSize, bold: true, color: theme.textColor,
          });
          yPos += 0.5;
          break;
        }

        case 'paragraph': {
          const rawText = token.text || '';
          if (/[├└│──]/.test(rawText)) {
            const treeLines = rawText.split('\n');
            const treeH = Math.max(0.4, treeLines.length * 0.22 + 0.1);
            const treeSegments = [];
            treeLines.forEach((line, idx) => {
              if (idx > 0) treeSegments.push({ text: '\n' });
              const lineSegs = parseInlineForPptx(line, {
                fontSize: 11, fontFace: 'Courier New', color: theme.textColor,
              });
              treeSegments.push(...lineSegs);
            });
            slide.addText(treeSegments, {
              x: MARGIN, y: yPos, w: CONTENT_W, h: treeH,
              fontSize: 11, fontFace: 'Courier New', color: theme.textColor, valign: 'top',
            });
            yPos += treeH + 0.1;
            break;
          }

          const segments = parseInlineForPptx(rawText, { fontSize: 14, color: theme.textColor });
          const lineCount = Math.ceil(rawText.length / 90);
          const boxH = Math.max(0.35, lineCount * 0.25);
          slide.addText(segments, {
            x: MARGIN, y: yPos, w: CONTENT_W, h: boxH,
            fontSize: 14, color: theme.textColor, valign: 'top',
          });
          yPos += boxH + 0.1;
          break;
        }

        case 'list': {
          const listItems = token.items || [];
          const listSegments = [];
          for (let idx = 0; idx < listItems.length; idx++) {
            const item = listItems[idx];
            const itemSegs = parseInlineForPptx(item.text || '', {
              fontSize: 13, color: theme.textColor,
            });
            if (itemSegs.length > 0) {
              itemSegs[0] = {
                text: itemSegs[0].text,
                options: {
                  ...itemSegs[0].options,
                  bullet: token.ordered ? { type: 'number' } : { characterCode: '2022' },
                  ...(idx > 0 ? { breakLine: true } : {}),
                },
              };
            }
            listSegments.push(...itemSegs);
          }
          const listH = Math.max(0.5, listItems.length * 0.28 + 0.1);
          slide.addText(listSegments, {
            x: INDENT_X, y: yPos, w: INDENT_W, h: listH,
            fontSize: 13, color: theme.textColor, valign: 'top',
          });
          yPos += listH + 0.1;
          break;
        }

        case 'blockquote': {
          const quoteText = (token.tokens?.map((t) => t.text || t.raw || '').join(' ')) || token.text || '';
          const quoteLines = Math.ceil(quoteText.length / 85);
          const quoteH = Math.max(0.5, quoteLines * 0.25);
          slide.addText('', {
            x: MARGIN, y: yPos, w: 0.08, h: quoteH,
            fill: { color: theme.accentColor },
          });
          const quoteSegs = parseInlineForPptx(quoteText, {
            fontSize: 13, italic: true, color: theme.quoteColor,
          });
          slide.addText([
            { text: '\u201C', options: { fontSize: 13, italic: true, color: theme.quoteColor } },
            ...quoteSegs,
            { text: '\u201D', options: { fontSize: 13, italic: true, color: theme.quoteColor } },
          ], {
            x: INDENT_X, y: yPos, w: INDENT_W, h: quoteH,
            fontSize: 13, italic: true, color: theme.quoteColor, valign: 'top',
          });
          yPos += quoteH + 0.1;
          break;
        }

        case 'table': {
          if (token.header && token.rows) {
            const colCount = token.header.length;
            const tableW = Math.min(CONTENT_W, Math.max(colCount * 2.0, 6));
            const headerRow = token.header.map((cell) => ({
              text: parseInlineForPptx(cell.text, {
                bold: true, fontSize: 11, color: theme.tableHeaderColor,
              }),
              options: { fill: { color: theme.tableHeaderBg } },
            }));
            const dataRows = token.rows.map((row) =>
              row.map((cell) => ({
                text: parseInlineForPptx(cell.text, { fontSize: 11, color: theme.textColor }),
                options: {},
              }))
            );
            const tableData = [headerRow, ...dataRows];
            const rowCount = tableData.length;
            slide.addTable(tableData, {
              x: MARGIN, y: yPos, w: tableW,
              fontSize: 11, color: theme.textColor,
              border: { type: 'solid', pt: 1, color: theme.tableBorderColor },
              colW: Array(colCount).fill(tableW / colCount),
            });
            yPos += 0.35 * rowCount + 0.2;
          }
          break;
        }

        case 'code': {
          const codeText = token.text || '';
          const mermaidImg = token.lang === 'mermaid' && codeText && mermaidImages?.get(codeText);
          if (mermaidImg) {
            const maxW = CONTENT_W;
            const maxH = 6.8 - yPos;
            const aspectRatio = mermaidImg.width / mermaidImg.height;
            let imgW = maxW;
            let imgH = imgW / aspectRatio;
            if (imgH > maxH) {
              imgH = maxH;
              imgW = imgH * aspectRatio;
            }
            imgW = Math.min(imgW, 10);
            imgH = Math.min(imgH, 5);
            slide.addImage({
              data: mermaidImg.dataUrl,
              x: MARGIN + (CONTENT_W - imgW) / 2,
              y: yPos, w: imgW, h: imgH,
            });
            yPos += imgH + 0.2;
            break;
          }

          const codeLineCount = codeText.split('\n').length;
          const codeH = Math.max(0.4, codeLineCount * 0.2 + 0.15);
          slide.addText(codeText, {
            x: MARGIN, y: yPos, w: CONTENT_W, h: codeH,
            fontSize: 10, fontFace: 'Courier New', color: theme.textColor,
            fill: { color: theme.bgColor === 'FFFFFF' || theme.bgColor === 'F8FAFC' ? 'F3F4F6' : '0D1117' },
            valign: 'top',
          });
          yPos += codeH + 0.1;
          break;
        }

        case 'hr':
        case 'space':
          yPos += 0.15;
          break;

        default:
          if (token.text || token.raw) {
            const defaultSegs = parseInlineForPptx(token.text || token.raw, {
              fontSize: 13, color: theme.textColor,
            });
            slide.addText(defaultSegs, {
              x: MARGIN, y: yPos, w: CONTENT_W, h: 0.35,
              fontSize: 13, color: theme.textColor, valign: 'top',
            });
            yPos += 0.4;
          }
      }
    }
  }

  // Convert {key: {dataUrl, width, height}} (from app.js captureMermaidImages)
  // into the Map<source, {dataUrl, width, height}> shape addContentToSlide expects.
  function imagesObjectToMap(imagesObj) {
    const map = new Map();
    if (!imagesObj) return map;
    for (const [k, v] of Object.entries(imagesObj)) {
      if (v && v.dataUrl) map.set(k, v);
    }
    return map;
  }

  async function exportToPptx({ title, content, theme: themeName = 'light', author, mermaidImages }) {
    const Pptx = getPptxGenJS();
    const markedLib = await loadMarked();
    const theme = THEMES[themeName] || THEMES.light;
    const slides = window.LumaSlides.parseSlides(content);
    const mermaidMap = imagesObjectToMap(mermaidImages);

    const pptx = new Pptx();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.title = title;
    if (author) pptx.author = author;

    for (let i = 0; i < slides.length; i++) {
      const slideData = slides[i];
      const slide = pptx.addSlide();
      const isFirstSlide = i === 0;

      slide.background = { fill: isFirstSlide ? theme.titleBg : theme.bgColor };

      if (isFirstSlide) {
        slide.addText(slideData.title, {
          x: MARGIN, y: 2.0, w: CONTENT_W, h: 0.8,
          fontSize: 36, bold: true, color: theme.titleColor,
          align: 'center', valign: 'middle',
        });

        let contentStartY = 3.0;
        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: MARGIN, y: 3.0, w: CONTENT_W, h: 0.5,
            fontSize: 20,
            color: themeName === 'branded' ? 'E0E0FF' : theme.subtitleColor,
            align: 'center', valign: 'top',
          });
          contentStartY = 3.7;
        }

        if (slideData.content) {
          addContentToSlide(slide, slideData.content, {
            ...theme,
            textColor: themeName === 'branded' ? 'E0E0E0' : theme.textColor,
          }, contentStartY, mermaidMap, markedLib);
        }
      } else {
        slide.addText(slideData.title, {
          x: MARGIN, y: 0.3, w: CONTENT_W, h: 0.7,
          fontSize: 24, bold: true, color: theme.contentTitleColor,
          valign: 'middle',
        });

        slide.addText('', {
          x: MARGIN, y: 1.1, w: 1.5, h: 0.05,
          fill: { color: theme.accentColor },
        });

        let contentStartY = 1.3;
        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: MARGIN, y: 1.25, w: CONTENT_W, h: 0.4,
            fontSize: 16, color: theme.subtitleColor, valign: 'top',
          });
          contentStartY = 1.75;
        }

        if (slideData.content) {
          addContentToSlide(slide, slideData.content, theme, contentStartY, mermaidMap, markedLib);
        }

        slide.addText(`${i + 1}`, {
          x: SLIDE_W - 1.1, y: 6.9, w: 0.5, h: 0.3,
          fontSize: 9, color: theme.accentColor, align: 'right',
        });
      }
    }

    // Output as ArrayBuffer so main process can write it via dialog
    return await pptx.write({ outputType: 'arraybuffer' });
  }

  window.LumaPpt = { exportToPptx, THEMES };
})();
