const MarkdownIt = require('markdown-it');
const hljs = require('highlight.js');

/**
 * Create a markdown-it instance configured for SharePoint
 * - Uses inline styles for better compatibility
 * - Supports syntax highlighting for code blocks
 * - Preserves all markdown formatting elements
 */
function createSharePointMarkdownConverter() {
  const md = new MarkdownIt({
    html: true,
    linkify: true,
    typographer: true,
    highlight: function (str, lang) {
      if (lang && hljs.getLanguage(lang)) {
        try {
          const highlighted = hljs.highlight(str, { language: lang, ignoreIllegals: true }).value;
          return `<pre class="ms-rteElement-CodeHTML" style="border: 1px solid #666666; background-color: #f5f5f5; padding: 12px; margin: 10px 0; overflow-x: auto; font-family: 'Consolas', 'Monaco', 'Courier New', monospace; font-size: 13px; line-height: 1.5;"><code class="language-${lang}" style="background-color: transparent; padding: 0; border: none; font-family: inherit;">${highlighted}</code></pre>`;
        } catch (__) {}
      }
      return `<pre class="ms-rteElement-CodeHTML" style="border: 1px solid #666666; background-color: #f5f5f5; padding: 12px; margin: 10px 0; overflow-x: auto; font-family: 'Consolas', 'Monaco', 'Courier New', monospace; font-size: 13px; line-height: 1.5;"><code style="background-color: transparent; padding: 0; border: none; font-family: inherit;">${md.utils.escapeHtml(str)}</code></pre>`;
    }
  });

  return md;
}

/**
 * Convert markdown string to HTML with SharePoint-compatible styling
 * @param {string} markdown - The markdown content to convert
 * @returns {string} HTML string ready for SharePoint
 */
function markdownToSharePointHTML(markdown) {
  const md = createSharePointMarkdownConverter();
  let html = md.render(markdown);
  html = addInlineStyles(html);
  return html;
}

/**
 * Add inline styles to HTML elements for SharePoint compatibility
 * This is needed because SharePoint strips external stylesheets and some classes
 */
function addInlineStyles(html) {
  const styleRules = [
    {
      pattern: /<h1>/g,
      replacement: '<h1 style="font-size: 2em; font-weight: bold; margin: 0.67em 0; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em;">'
    },
    {
      pattern: /<h2>/g,
      replacement: '<h2 style="font-size: 1.5em; font-weight: bold; margin: 0.83em 0; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em;">'
    },
    {
      pattern: /<h3>/g,
      replacement: '<h3 style="font-size: 1.17em; font-weight: bold; margin: 1em 0;">'
    },
    {
      pattern: /<h4>/g,
      replacement: '<h4 style="font-size: 1em; font-weight: bold; margin: 1.33em 0;">'
    },
    {
      pattern: /<h5>/g,
      replacement: '<h5 style="font-size: 0.83em; font-weight: bold; margin: 1.67em 0;">'
    },
    {
      pattern: /<h6>/g,
      replacement: '<h6 style="font-size: 0.67em; font-weight: bold; margin: 2.33em 0;">'
    },
    {
      pattern: /<p>/g,
      replacement: '<p style="margin: 1em 0; line-height: 1.6;">'
    },
    {
      pattern: /<strong>/g,
      replacement: '<strong style="font-weight: bold;">'
    },
    {
      pattern: /<b>/g,
      replacement: '<b style="font-weight: bold;">'
    },
    {
      pattern: /<em>/g,
      replacement: '<em style="font-style: italic;">'
    },
    {
      pattern: /<i>/g,
      replacement: '<i style="font-style: italic;">'
    },
    {
      pattern: /<a([^>]*?)>/g,
      replacement: '<a$1 style="color: #0366d6; text-decoration: none;">'
    },
    {
      pattern: /<ul>/g,
      replacement: '<ul style="padding-left: 2em; margin: 1em 0;">'
    },
    {
      pattern: /<ol>/g,
      replacement: '<ol style="padding-left: 2em; margin: 1em 0;">'
    },
    {
      pattern: /<li>/g,
      replacement: '<li style="margin: 0.5em 0;">'
    },
    {
      pattern: /<blockquote>/g,
      replacement: '<blockquote style="border-left: 4px solid #dfe2e5; padding: 0 1em; color: #6a737d; margin: 1em 0;">'
    },
    {
      pattern: /<table>/g,
      replacement: '<table style="border-collapse: collapse; width: 100%; margin: 1em 0; border-spacing: 0;">'
    },
    {
      pattern: /<th>/g,
      replacement: '<th style="border: 1px solid #dfe2e5; padding: 6px 13px; font-weight: bold; background-color: #f6f8fa;">'
    },
    {
      pattern: /<td>/g,
      replacement: '<td style="border: 1px solid #dfe2e5; padding: 6px 13px;">'
    },
    {
      pattern: /<hr>/g,
      replacement: '<hr style="height: 0.25em; padding: 0; margin: 24px 0; background-color: #e1e4e8; border: 0;">'
    },
    {
      pattern: /<img([^>]*?)>/g,
      replacement: '<img$1 style="max-width: 100%; height: auto; display: block; margin: 1em 0;">'
    },
    {
      pattern: /<code>/g,
      replacement: '<code style="background-color: rgba(27, 31, 35, 0.05); padding: 0.2em 0.4em; margin: 0; font-size: 85%; border-radius: 3px; font-family: Consolas, Monaco, \'Courier New\', monospace;">'
    }
  ];

  let styledHTML = html;
  styleRules.forEach(rule => {
    styledHTML = styledHTML.replace(rule.pattern, rule.replacement);
  });

  return styledHTML;
}

module.exports = {
  createSharePointMarkdownConverter,
  markdownToSharePointHTML,
  addInlineStyles
};
