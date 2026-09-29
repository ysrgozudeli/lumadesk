// Slide parser shared by PPT export and Presentation mode.
// Convention 0 (Marp): frontmatter "marp: true" or "---" slide separators
// Convention 1: "# SLIDE N -- Title" markers
// Convention 2 (fallback): split on h1 headings, or h2 if only one h1
//
// Exposes:
//   window.LumaSlides.parseSlides(markdown) -> [{ title, subtitle?, content, sectionClass? }]
//   window.LumaSlides.parseDeck(markdown)   -> { slides, style }
(function () {
  function cleanBreaks(text) {
    return text.replace(/^(\*{3,}|-{3,})$/gm, '').trim();
  }

  function stripFrontmatter(md) {
    return md.replace(/^﻿?---[ \t]*\r?\n[\s\S]*?\r?\n---[ \t]*(\r?\n|$)/, '');
  }

  // A deck is Marp/slide-separated when it declares marp:true in frontmatter
  // or uses standalone "---" lines as slide breaks.
  function isMarpDeck(md) {
    const fm = md.match(/^﻿?---[ \t]*\r?\n([\s\S]*?)\r?\n---[ \t]*(\r?\n|$)/);
    if (fm && /^\s*marp\s*:\s*true\s*$/m.test(fm[1])) return true;
    // Count standalone "---" separators in the body (after any frontmatter).
    const body = stripFrontmatter(md);
    const seps = body.match(/^[ \t]*---[ \t]*$/gm);
    return !!seps && seps.length >= 1;
  }

  // Split a Marp deck on "---" separators. Each chunk becomes one slide,
  // keeping its markdown body intact (code blocks, tables, HTML spans).
  function parseMarpDeck(markdown) {
    const noFm = stripFrontmatter(markdown);

    // Pull out any <style> block(s) so they can be applied to the whole deck
    // instead of being rendered as literal CSS text on the first slide.
    let style = '';
    const body = noFm.replace(/<style[\s\S]*?<\/style>/gi, (m) => {
      style += m.replace(/^<style[^>]*>/i, '').replace(/<\/style>\s*$/i, '') + '\n';
      return '';
    });

    const chunks = body.split(/^[ \t]*---[ \t]*$/m);
    const slides = [];

    for (const rawChunk of chunks) {
      // Capture a Marp per-slide directive (<!-- _class: lead -->) then drop
      // all HTML comments from the content.
      const classMatch = rawChunk.match(/<!--\s*_class:\s*([^>]+?)\s*-->/);
      const sectionClass = classMatch ? classMatch[1].trim() : undefined;
      const chunk = rawChunk.replace(/<!--[\s\S]*?-->/g, '').trim();
      if (!chunk) continue;

      const lines = chunk.split('\n');
      let title = '';
      let subtitle;
      let titleTaken = false;
      let subtitleTaken = false;
      const bodyLines = [];

      for (const line of lines) {
        const h1 = line.match(/^#\s+(.+)$/);
        const h2 = line.match(/^##\s+(.+)$/);
        if (!titleTaken && h1) { title = h1[1].trim(); titleTaken = true; continue; }
        if (!titleTaken && h2) { title = h2[1].trim(); titleTaken = true; continue; }
        if (titleTaken && !subtitleTaken && !bodyLines.some((l) => l.trim()) && h2) {
          subtitle = h2[1].trim(); subtitleTaken = true; continue;
        }
        bodyLines.push(line);
      }

      slides.push({
        title: title || 'Slide',
        subtitle,
        content: bodyLines.join('\n').trim(),
        sectionClass,
      });
    }

    return { slides, style };
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
    // Marp / "---"-separated decks: honour the author's slide boundaries.
    if (isMarpDeck(markdown)) {
      return parseMarpDeck(markdown).slides;
    }

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

  // Full deck info (slides + extracted <style>) for Presentation mode.
  function parseDeck(markdown) {
    if (isMarpDeck(markdown)) return parseMarpDeck(markdown);
    return { slides: parseSlides(markdown), style: '' };
  }

  window.LumaSlides = { parseSlides, parseSlidesFromHeadings, parseDeck };
})();
