# LumaDesk

A lightweight desktop app for previewing Markdown documents and exporting them to Word. Built with Electron.

## Features

- **Folder browsing** — open any folder and browse `.md` / `.txt` files in a sidebar tree
- **Live preview** — clean, styled Markdown rendering with auto-refresh on file changes
- **Mermaid diagrams** — full-color diagram rendering with click-to-expand fullscreen modal (zoom & pan)
- **Word export** — export to `.docx` with proper headings, tables, code blocks, and embedded Mermaid diagrams
- **Auto-generated Table of Contents** — `# Table of Contents` headings become a real Word TOC
- **Pandoc heading IDs** — `{#custom-id}` syntax is stripped from display and used as bookmarks
- **File type filter** — configurable extensions (default: `.md`, `.txt`)
- **Resizable sidebar** — drag to resize
- **Keyboard shortcuts** — `Ctrl+O` open folder, `Ctrl+E` export Word, `Esc` close modal

## Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) 18+

### Install

```bash
git clone https://github.com/peerluma/lumadesk.git
cd lumadesk
npm install
```

### Run

```bash
npm start
```

### Build for Windows

```bash
# Installer + portable
npm run dist:win

# Portable only
npm run dist:portable
```

Output goes to `dist/`.

## Usage

1. Click **Open Folder** (or `Ctrl+O`) and select a folder with Markdown files
2. Browse files in the sidebar — folders become collapsible sections
3. Click a file to preview it
4. Click **Export Word** (or `Ctrl+E`) to save as `.docx`
5. Click any Mermaid diagram to view it fullscreen with zoom and pan

### Folder Structure

Numbered prefixes control sort order but are stripped from display:

```
my-docs/
  01-Getting Started/
    intro.md
    setup.md
  02-Architecture/
    overview.md
```

Shows as:

```
Getting Started
  intro
  setup
Architecture
  overview
```

## Tech Stack

- [Electron](https://www.electronjs.org/) — desktop shell
- [marked](https://marked.js.org/) — Markdown parsing
- [mermaid](https://mermaid.js.org/) — diagram rendering
- [docx](https://docx.js.org/) — Word document generation
- [chokidar](https://github.com/paulmillr/chokidar) — file watching

## License

[MIT](LICENSE)
