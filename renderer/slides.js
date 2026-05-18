// Slide parser shared by PPT export and Presentation mode.
// Convention 1 (preferred): "# SLIDE N -- Title" markers
// Convention 2 (fallback): split on h1 headings, or h2 if only one h1
//
// Exposes: window.LumaSlides.parseSlides(markdown) -> [{ title, subtitle?, content }]
(function () {
  function cleanBreaks(text) {
    return text.replace(/^(\*{3,}|-{3,})$/gm, '').trim();
  }

  function parseSlidesFromHeadings(markdown) {
    const h1Matches = [...markdown.matchAll(/^#\s+(.+)$/gm)];
    const h2Matches = [...markdown.matchAll(/^##\s+(.+)$/gm)];

    const useBothLevels = h1Matches.length >= 2 && h2Matches.length > h1Matches.length;

    if (useBothLevels) {
      const allHeadings = [];
      for (const m of markdown.matchAll(/^(#{1,2})\s+(.+)$/gm)) {
        const level = m[1].length;
        if (level <= 2) {
          allHeadings.push({ title: m[2].trim(), index: m.index, length: m[0].length, level });
        }
      }
      allHeadings.sort((a, b) => a.index - b.index);

      const slides = [];
      const preamble = cleanBreaks(markdown.slice(0, allHeadings[0]?.index ?? 0));
      if (preamble) {
        slides.push({ title: preamble.split('\n')[0] || 'Introduction', content: '' });
      }

      for (let i = 0; i < allHeadings.length; i++) {
        const h = allHeadings[i];
        const contentStart = h.index + h.length;
        const contentEnd = i < allHeadings.length - 1 ? allHeadings[i + 1].index : markdown.length;
        const content = cleanBreaks(markdown.slice(contentStart, contentEnd));

        if (h.level === 1) {
          slides.push({ title: h.title, content });
        } else {
          const subMatch = content.match(/^###\s+(.+)$/m);
          const subtitle = subMatch ? subMatch[1].trim() : undefined;
          const body = subtitle ? content.replace(/^###\s+.+$/m, '').trim() : content;
          slides.push({ title: h.title, subtitle, content: body });
        }
      }

      return slides;
    }

    const useH2 = h1Matches.length <= 1 && h2Matches.length >= 2;
    const splitRegex = useH2 ? /^##\s+(.+)$/gm : /^#\s+(.+)$/gm;

    const headings = [];
    let match;
    while ((match = splitRegex.exec(markdown)) !== null) {
      headings.push({ title: match[1].trim(), index: match.index, length: match[0].length });
    }

    if (headings.length === 0) {
      return [{ title: 'Presentation', content: markdown.trim() }];
    }

    const slides = [];
    const preamble = cleanBreaks(markdown.slice(0, headings[0].index));
    if (preamble) {
      const h1InPreamble = preamble.match(/^#\s+(.+)$/m);
      if (h1InPreamble) {
        const titleContent = cleanBreaks(preamble.replace(/^#\s+.+$/m, ''));
        slides.push({ title: h1InPreamble[1].trim(), content: titleContent });
      } else {
        slides.push({ title: 'Introduction', content: preamble });
      }
    }

    for (let i = 0; i < headings.length; i++) {
      const contentStart = headings[i].index + headings[i].length;
      const contentEnd = i < headings.length - 1 ? headings[i + 1].index : markdown.length;
      const content = cleanBreaks(markdown.slice(contentStart, contentEnd));

      const subRegex = useH2 ? /^###\s+(.+)$/m : /^##\s+(.+)$/m;
      let subtitle;
      const subMatch = content.match(subRegex);
      if (subMatch) subtitle = subMatch[1].trim();

      const contentWithoutSubtitle = subtitle ? content.replace(subRegex, '').trim() : content;

      slides.push({ title: headings[i].title, subtitle, content: contentWithoutSubtitle });
    }

    return slides;
  }

  function parseSlides(markdown) {
    const slideRegex = /^#\s+SLIDE\s+\d+\s*--\s*(.+)$/gm;
    const matches = [];
    let match;
    while ((match = slideRegex.exec(markdown)) !== null) {
      matches.push({ title: match[1].trim(), index: match.index });
    }

    if (matches.length === 0) {
      return parseSlidesFromHeadings(markdown);
    }

    const slides = [];

    const preamble = markdown.slice(0, matches[0].index).replace(/^---$/gm, '').trim();
    if (preamble) {
      const titleMatch = preamble.match(/^#\s+(.+)$/m);
      const preambleTitle = titleMatch ? titleMatch[1].trim() : 'Cover';
      let preambleContent = titleMatch ? preamble.replace(/^#\s+.+$/m, '').trim() : preamble;

      let preambleSubtitle;
      const subMatch = preambleContent.match(/^##\s+(.+)$/m);
      if (subMatch) {
        preambleSubtitle = subMatch[1].trim();
        preambleContent = preambleContent.replace(/^##\s+.+$/m, '').trim();
      }

      slides.push({ title: preambleTitle, subtitle: preambleSubtitle, content: preambleContent });
    }

    for (let i = 0; i < matches.length; i++) {
      const start = matches[i].index;
      const end = i < matches.length - 1 ? matches[i + 1].index : markdown.length;
      const slideBlock = markdown.slice(start, end);

      const contentAfterHeading = slideBlock.replace(/^#\s+SLIDE\s+\d+\s*--\s*.+$/m, '').trim();
      const cleanedContent = contentAfterHeading.replace(/^---$/gm, '').trim();

      let subtitle;
      const subtitleMatch = cleanedContent.match(/^##\s+(.+)$/m);
      if (subtitleMatch) subtitle = subtitleMatch[1].trim();

      const contentWithoutSubtitle = cleanedContent.replace(/^##\s+.+$/m, '').trim();

      slides.push({ title: matches[i].title, subtitle, content: contentWithoutSubtitle });
    }

    return slides;
  }

  window.LumaSlides = { parseSlides, parseSlidesFromHeadings };
})();
