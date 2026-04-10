const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  Table,
  TableRow,
  TableCell,
  TableOfContents,
  BookmarkStart,
  BookmarkEnd,
  ImageRun,
} = require('docx');
const fs = require('fs');

let marked;
async function getMarked() {
  if (!marked) {
    const mod = await import('marked');
    marked = mod.marked;
  }
  return marked;
}

/**
 * Export markdown content to a Word document.
 */
async function exportToWord({ title, content, author, savePath, mermaidImages }) {
  const contentParagraphs = await markdownToDocxParagraphs(content, mermaidImages);

  // Header
  const headerParagraphs = [
    new Paragraph({
      children: [new TextRun({ text: title, bold: true, size: 48, color: '1F2937' })],
      spacing: { after: 120 },
    }),
  ];

  if (author) {
    headerParagraphs.push(new Paragraph({
      children: [new TextRun({ text: `By ${author}`, color: '6B7280', size: 20 })],
      spacing: { after: 60 },
    }));
  }

  headerParagraphs.push(new Paragraph({
    children: [new TextRun({ text: `Exported ${new Date().toLocaleDateString()}`, color: '9CA3AF', size: 18 })],
    spacing: { after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'E5E7EB' } },
  }));

  const doc = new Document({
    features: { updateFields: true },
    sections: [{
      children: [
        ...headerParagraphs,
        ...contentParagraphs,
        new Paragraph({
          children: [new TextRun({ text: 'Exported from LumaDesk', color: '9CA3AF', size: 16 })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 480 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'E5E7EB' } },
        }),
      ],
    }],
    numbering: {
      config: [{
        reference: 'default-numbering',
        levels: [{
          level: 0,
          format: 'decimal',
          text: '%1.',
          alignment: AlignmentType.LEFT,
        }],
      }],
    },
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(savePath, buffer);
}

// ---- Markdown → docx conversion ----

function stripPandocId(text) {
  return text.replace(/\s*\{#[^}]+\}\s*$/, '').trim();
}

function extractPandocId(text) {
  const match = text.match(/\{#([^}]+)\}\s*$/);
  return match ? match[1] : null;
}

function headingToBookmarkId(text) {
  return stripPandocId(text).toLowerCase().replace(/[^\w\s-]/g, '').replace(/\s+/g, '-');
}

function dataUrlToBuffer(dataUrl) {
  const base64 = dataUrl.split(',')[1];
  const binary = Buffer.from(base64, 'base64');
  return binary;
}

async function markdownToDocxParagraphs(markdown, mermaidImages) {
  const m = await getMarked();
  const paragraphs = [];
  const tokens = m.lexer(markdown);
  let bookmarkCounter = 1;
  let skipTocLinks = false;

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    switch (token.type) {
      case 'heading': {
        const headingLevel = token.depth === 1 ? HeadingLevel.HEADING_1 :
                            token.depth === 2 ? HeadingLevel.HEADING_2 :
                            token.depth === 3 ? HeadingLevel.HEADING_3 :
                            HeadingLevel.HEADING_4;
        const rawText = token.text || '';
        const pandocId = extractPandocId(rawText);
        const cleanText = stripPandocId(rawText);

        // Detect "Table of Contents" heading → insert real Word TOC
        if (/^table\s+of\s+contents$/i.test(cleanText)) {
          paragraphs.push(new Paragraph({
            children: [new TextRun({ text: cleanText })],
            heading: headingLevel,
            spacing: { before: 240, after: 120 },
          }));
          paragraphs.push(new TableOfContents("Table of Contents", {
            hyperlink: true,
            headingStyleRange: "1-4",
          }));
          skipTocLinks = true;
          break;
        }

        skipTocLinks = false;
        const bookmarkName = pandocId || headingToBookmarkId(cleanText);
        const numericId = bookmarkCounter++;

        paragraphs.push(new Paragraph({
          children: [
            new BookmarkStart(bookmarkName, numericId),
            new TextRun({ text: cleanText }),
            new BookmarkEnd(numericId),
          ],
          heading: headingLevel,
          spacing: { before: 240, after: 120 },
        }));
        break;
      }

      case 'paragraph': {
        // Skip TOC anchor link paragraphs
        if (skipTocLinks) {
          const innerTokens = token.tokens || [];
          const isAnchorLink = innerTokens.length === 1 && innerTokens[0].type === 'link'
            && (innerTokens[0].href || '').startsWith('#');
          if (isAnchorLink) break;
          skipTocLinks = false;
        }
        paragraphs.push(new Paragraph({
          children: parseInlineTokens(token.tokens || [], token.text || ''),
          spacing: { before: 120, after: 120 },
        }));
        break;
      }

      case 'list': {
        const items = token.items || [];
        items.forEach((item) => {
          paragraphs.push(new Paragraph({
            children: parseInlineTokens(item.tokens || [], item.text || ''),
            bullet: token.ordered ? undefined : { level: 0 },
            numbering: token.ordered ? { reference: 'default-numbering', level: 0 } : undefined,
            spacing: { before: 60, after: 60 },
          }));
        });
        break;
      }

      case 'code': {
        // Mermaid diagrams → embed as image
        if (token.lang === 'mermaid' && mermaidImages) {
          const imgData = mermaidImages[token.text];
          if (imgData) {
            try {
              const buffer = dataUrlToBuffer(imgData.dataUrl);
              const maxWidth = 600;
              const scale = imgData.width > maxWidth ? maxWidth / imgData.width : 1;
              paragraphs.push(new Paragraph({
                children: [new ImageRun({
                  data: buffer,
                  transformation: {
                    width: Math.round(imgData.width * scale),
                    height: Math.round(imgData.height * scale),
                  },
                  type: 'png',
                })],
                spacing: { before: 120, after: 120 },
              }));
            } catch (e) {
              // Fallback to code block on error
              paragraphs.push(new Paragraph({
                children: [new TextRun({ text: '[Mermaid diagram]', italics: true, color: '6B7280' })],
              }));
            }
            break;
          }
        }

        const lines = (token.text || '').split('\n');
        for (const line of lines) {
          paragraphs.push(new Paragraph({
            children: [new TextRun({
              text: line || ' ',
              font: 'Consolas',
              size: 18,
              color: 'F9FAFB',
            })],
            shading: { fill: '1F2937' },
            spacing: { before: 0, after: 0 },
          }));
        }
        break;
      }

      case 'blockquote': {
        const innerTokens = token.tokens || [];
        for (const inner of innerTokens) {
          if (inner.type === 'paragraph') {
            paragraphs.push(new Paragraph({
              children: parseInlineTokens(inner.tokens || [], inner.text || ''),
              spacing: { before: 60, after: 60 },
              indent: { left: 720 },
              border: { left: { style: BorderStyle.SINGLE, size: 12, color: 'D1D5DB' } },
            }));
          }
        }
        break;
      }

      case 'table': {
        if (token.header && token.rows) {
          const colCount = token.header.length;
          const cellWidth = Math.floor(9360 / colCount);
          const tableRows = [];

          // Header row
          tableRows.push(new TableRow({
            children: token.header.map(cell => new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ text: cell.text || '', bold: true })],
              })],
              shading: { fill: 'F3F4F6' },
              width: { size: cellWidth, type: 'dxa' },
            })),
          }));

          // Data rows
          for (const row of token.rows) {
            tableRows.push(new TableRow({
              children: row.map(cell => new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ text: cell.text || '' })],
                })],
                width: { size: cellWidth, type: 'dxa' },
              })),
            }));
          }

          paragraphs.push(new Table({
            rows: tableRows,
            width: { size: 9360, type: 'dxa' },
          }));
        }
        break;
      }

      case 'hr':
        paragraphs.push(new Paragraph({
          children: [],
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'E5E7EB' } },
          spacing: { before: 240, after: 240 },
        }));
        break;

      case 'space':
        break;

      default:
        if (token.raw && token.raw.trim()) {
          paragraphs.push(new Paragraph({
            children: [new TextRun({ text: token.raw.trim() })],
            spacing: { before: 120, after: 120 },
          }));
        }
    }
  }

  return paragraphs;
}

/**
 * Convert marked inline tokens to docx TextRuns.
 */
function parseInlineTokens(tokens, fallbackText) {
  if (tokens && tokens.length > 0) {
    const runs = collectRuns(tokens);
    if (runs.length > 0) return runs;
  }
  return [new TextRun({ text: fallbackText || '' })];
}

function collectRuns(tokens, style = {}) {
  const runs = [];

  for (const t of tokens) {
    switch (t.type) {
      case 'strong':
        if (t.tokens && t.tokens.length > 0) {
          runs.push(...collectRuns(t.tokens, { ...style, bold: true }));
        } else {
          runs.push(new TextRun({ text: t.text || '', ...style, bold: true }));
        }
        break;
      case 'em':
        if (t.tokens && t.tokens.length > 0) {
          runs.push(...collectRuns(t.tokens, { ...style, italics: true }));
        } else {
          runs.push(new TextRun({ text: t.text || '', ...style, italics: true }));
        }
        break;
      case 'codespan':
        runs.push(new TextRun({
          text: t.text || '', ...style,
          font: 'Consolas',
          shading: { fill: 'E5E7EB', type: 'clear', color: 'auto' },
        }));
        break;
      case 'link':
        runs.push(new TextRun({
          text: t.text || t.href || '', ...style,
          color: '2563EB', underline: {},
        }));
        break;
      case 'text':
        if (t.tokens && t.tokens.length > 0) {
          runs.push(...collectRuns(t.tokens, style));
        } else {
          runs.push(new TextRun({ text: t.text || t.raw || '', ...style }));
        }
        break;
      default:
        runs.push(new TextRun({ text: t.raw || t.text || '', ...style }));
    }
  }

  return runs;
}

module.exports = { exportToWord };
