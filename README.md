# Markdown to SharePoint HTML Converter

A Node.js CLI tool that converts Markdown files to HTML optimized for SharePoint with syntax-highlighted code blocks and inline styles.

## Features

- **Full Markdown Support**: Headings, bold/italic, lists, tables, links, images, blockquotes
- **Syntax Highlighting**: Code blocks with highlight.js for 190+ languages
- **Table of Contents**: Automatic TOC generation with clickable anchor links
- **Anchor Links**: All headings have anchor IDs for internal page navigation
- **Dark/Light Themes**: Theme support for SharePoint light or dark backgrounds
- **SharePoint-Optimized**: Inline styles (no external CSS) and SharePoint-compatible class names
- **CLI Tool**: Easy-to-use command-line interface
- **Preserves Formatting**: All markdown elements properly styled for SharePoint's rich text editor

## Installation

```bash
# Clone or copy this repository
cd markdown_to_sharepoint

# Install dependencies
npm install
```

## Usage

### Basic Usage

```bash
# Convert markdown to HTML and print to stdout
node bin/md2sp-html.js input.md

# Convert to a file
node bin/md2sp-html.js input.md output.html

# Convert with full HTML structure (includes <html>, <head>, <body>)
node bin/md2sp-html.js input.md output.html --wrap

# Convert with table of contents (requires [TOC] marker in markdown)
node bin/md2sp-html.js input.md output.html --toc

# Convert with custom TOC title
node bin/md2sp-html.js input.md output.html --toc --toc-title "Contents"

# Use dark theme for SharePoint with dark/black background
node bin/md2sp-html.js input.md output.html --dark-theme

# Dark theme with TOC
node bin/md2sp-html.js input.md output.html --dark-theme --toc

# Combine with other options
node bin/md2sp-html.js input.md output.html --toc --wrap

# Dark theme with all options
node bin/md2sp-html.js input.md output.html --dark-theme --toc --wrap
```

### Using the CLI

```bash
# View help
node bin/md2sp-html.js --help

# Convert a markdown document
node bin/md2sp-html.js README.md README-sharepoint.html --wrap
```

### Programmatic Usage

```javascript
const { markdownToSharePointHTML } = require('./index');

const markdown = `# Title

\`\`\`javascript
function hello() {
  console.log("Hello World");
}
\`\`\`
`;

const html = markdownToSharePointHTML(markdown);
console.log(html);
```

## SharePoint Integration

### Step 1: Convert Your Markdown

```bash
node bin/md2sp-html.js your-document.md output.html --wrap
```

### Step 2: Open Output File

Open the generated `output.html` file in a browser or text editor.

### Step 3: Copy to SharePoint

1. Open your SharePoint page in edit mode
2. Click on the rich text editor where you want to paste content
3. Use **Edit Source** (or **HTML Source**) option in the editor toolbar
4. Copy the content from the `<body>` tag of your HTML file (or the entire file if using --wrap)
5. Paste into the SharePoint HTML editor
6. Save and publish

## What Gets Preserved

| Markdown Element | HTML Output | Notes |
|-----------------|--------------|--------|
| Headings (H1-H6) | Styled headings with anchor IDs | Proper sizing, margins, borders, clickable navigation |
| **Bold** / *Italic* | Styled text | Font weight and style applied |
| `Inline code` | Styled code | Light background, monospace font |
| Code blocks (light theme) | Highlighted code | Light gray background (#f5f5f5) + **colored syntax** |
| Code blocks (dark theme) | Highlighted code | Dark gray background (#1e1e1e) + **colored syntax** |
| Ordered lists | Numbered lists | Proper nesting and spacing |
| Unordered lists | Bulleted lists | Proper nesting and spacing |
| [Links](url) | Styled links | Blue color, no underline, href preserved |
| [TOC links](#heading) | Anchor links | Clickable internal navigation |
| Images | Responsive images | Max-width, auto height |
| > Blockquotes (light) | Styled quotes | Light gray text, gray border |
| > Blockquotes (dark) | Styled quotes | Light gray text, dark border |
| Tables (light theme) | Styled tables | Light gray borders, header styling |
| Tables (dark theme) | Styled tables | Dark borders, dark header background |
| Horizontal rules | Styled HR | Proper height and spacing |
| Theme awareness | All elements | Adjusts colors based on --dark-theme flag |
| **Syntax highlighting** | Colored code | Inline span colors applied to highlighted code keywords, strings, numbers |

## Dark/Light Theme Support

For SharePoint sites with black or dark backgrounds, use to `--dark-theme` flag:

**Light Theme** (default):
- Code blocks: Light gray background (#f5f5f5) with **colored syntax highlighting**
- TOC: Light gray background (#f6f8fa)
- Optimized for light SharePoint backgrounds
- Syntax colors optimized for light mode (softer pastel tones)

**Dark Theme**:
- Code blocks: Dark gray background (#1e1e1e) with **colored syntax highlighting**
- TOC: Dark background (#2d3333)
- Better contrast for black/dark SharePoint backgrounds
- Syntax colors optimized for dark mode (brighter keywords/strings)

Example:
```bash
# Light theme (default)
node bin/md2sp-html.js input.md output.html

# Dark theme for black backgrounds
node bin/md2sp-html.js input.md output.html --dark-theme
```

## Table of Contents (TOC)

Add `[TOC]` marker in your markdown where you want the table of contents to appear:

```markdown
# My Document

[TOC]

## Section 1
Content here...

## Section 2
More content...
```

When converted with `--toc`, this will generate a clickable table of contents with links to each heading section.

## Supported Languages for Syntax Highlighting

Code blocks with language specifiers are automatically highlighted:

\`\`\`javascript
// JavaScript code
\`\`\`

\`\`\`python
# Python code
\`\`\`

\`\`\`bash
# Bash scripts
\`\`\`

\`\`\`powershell
# PowerShell scripts
\`\`\`

\`\`\`sql
# SQL queries
\`\`\`

\`\`\`yaml
# YAML configuration
\`\`\`

\`\`\`json
# JSON data
\`\`\`

Supported languages include: JavaScript, TypeScript, Python, Java, C#, PHP, Ruby, Go, Rust, Swift, Kotlin, SQL, Bash, PowerShell, HTML, CSS, JSON, XML, YAML, and 180+ more.

## SharePoint-Specific Features

### Inline Styles

All styles are applied inline to ensure SharePoint's content management system doesn't strip them:

```html
<h1 style="font-size: 2em; font-weight: bold;">Title</h1>
<p style="margin: 1em 0; line-height: 1.6;">Content</p>
```

### Code Block Classes

Code blocks use SharePoint's `ms-rteElement-CodeHTML` class for compatibility:

```html
<pre class="ms-rteElement-CodeHTML" style="...">
  <code class="language-javascript">...</code>
</pre>
```

### Why This Works in SharePoint

1. **Inline CSS**: SharePoint strips external stylesheets and `<style>` tags in the body
2. **No External Dependencies**: All styling is self-contained
3. **SharePoint Classes**: Uses `ms-rteElement-CodeHTML` for code blocks
4. **Proper HTML Structure**: Valid HTML5 that SharePoint accepts

## Limitations

- SharePoint may strip certain HTML attributes or classes
- JavaScript and iframes will be removed by SharePoint's security filters
- External stylesheets won't work (use inline styles)
- Some advanced markdown extensions may not be supported

## Testing

A test file is included to demonstrate all features:

```bash
# Run the built-in test
node bin/md2sp-html.js test.md test-output.html --wrap

# Open the result in your browser
open test-output.html
```

## Dependencies

- **markdown-it**: Markdown parser with full CommonMark support
- **markdown-it-anchor**: Adds anchor IDs to headings
- **markdown-it-toc-done-right**: Generates table of contents
- **highlight.js**: Syntax highlighting for code blocks
- **commander**: CLI framework

## License

MIT
