const MarkdownIt = require('markdown-it');
const hljs = require('highlight.js');
const anchor = require('markdown-it-anchor');
const toc = require('markdown-it-toc-done-right');

function createSharePointMarkdownConverter(options = {}) {
  const { includeTOC = false, tocTitle = 'Table of Contents', theme = 'light' } = options;

  const isDarkTheme = theme === 'dark';

  const codeBlockBackground = isDarkTheme ? '#1e1e1e' : '#f5f5f5';
  const codeBlockBorder = isDarkTheme ? '#4a5568' : '#666666';
  const codeBlockPadding = isDarkTheme ? '16px' : '12px';

  const md = new MarkdownIt({
    html: true,
    linkify: true,
    typographer: true,
    highlight: function (str, lang) {
      if (lang && hljs.getLanguage(lang)) {
        try {
          const highlighted = hljs.highlight(str, {
            language: lang,
            ignoreIllegals: true
          }).value;
          return `<pre class="ms-rteElement-CodeHTML" style="border: 1px solid ${codeBlockBorder}; background-color: ${codeBlockBackground}; padding: ${codeBlockPadding}; margin: 10px 0; overflow-x: auto; font-family: 'Consolas', 'Monaco', 'Courier New', monospace; font-size: 13px; line-height: 1.5;"><code class="language-${lang}" style="background-color: transparent; padding: 0; border: none; font-family: inherit;">${highlighted}</code></pre>`;
        } catch (__) {}
      }
      return `<pre class="ms-rteElement-CodeHTML" style="border: 1px solid ${codeBlockBorder}; background-color: ${codeBlockBackground}; padding: ${codeBlockPadding}; margin: 10px 0; overflow-x: auto; font-family: 'Consolas', 'Monaco', 'Courier New', monospace; font-size: 13px; line-height: 1.5;"><code style="background-color: transparent; padding: 0; border: none; font-family: inherit;">${md.utils.escapeHtml(str)}</code></pre>`;
    }
  });

  md.use(anchor, {
    permalink: false,
    slugify: (s) => s.toLowerCase().replace(/\s+/g, '-').replace(/[^\w-]+/g, ''),
    level: 1,
    tabIndex: false
  });

  if (includeTOC) {
    md.use(toc, {
      tocTitle: `<h2 style="font-size: 1.5em; font-weight: bold; margin: 0.83em 0; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em;">${tocTitle}</h2>`,
      tocClassName: 'markdown-it-toc',
      listType: 'ul',
      level: [1, 2, 3],
      listClassName: 'markdown-it-toc-list',
      itemClassName: 'markdown-it-toc-item',
      linkClassName: 'markdown-it-toc-link',
      callback: (html, ast) => {
        return html;
      }
    });
  }

  return md;
}

function markdownToSharePointHTML(markdown, options = {}) {
  const { includeTOC = false, tocTitle = 'Table of Contents', theme = 'light' } = options;
  const md = createSharePointMarkdownConverter({ includeTOC, tocTitle, theme });
  let html = md.render(markdown);
  html = addInlineStyles(html, theme);
  return html;
}

function addInlineStyles(html, theme = 'light') {
  const isDarkTheme = theme === 'dark';
  const tocBackground = isDarkTheme ? '#2d3333' : '#f6f8fa';
  const tocBorder = isDarkTheme ? '#4a5568' : '#dfe2e5';
  const blockquoteColor = isDarkTheme ? '#adb5bd' : '#6a737d';
  const blockquoteBorder = isDarkTheme ? '#4a5568' : '#dfe2e5';
  const hrColor = isDarkTheme ? '#4b5563' : '#e1e4e8';
  const headingBorder = isDarkTheme ? '#374151' : '#eaecef';
  const tableHeaderBackground = isDarkTheme ? '#161b22' : '#f6f8fa';
  const tableBorder = isDarkTheme ? '#4b5563' : '#dfe2e5';

  const styleRules = [
    {
      pattern: /<h1>/g,
      replacement: `<h1 style="font-size: 2em; font-weight: bold; margin: 0.67em 0; border-bottom: 1px solid ${headingBorder}; padding-bottom: 0.3em;">`
    },
    {
      pattern: /<h2>/g,
      replacement: `<h2 style="font-size: 1.5em; font-weight: bold; margin: 0.83em 0; border-bottom: 1px solid ${headingBorder}; padding-bottom: 0.3em;">`
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
      pattern: /<blockquote>/g,
      replacement: `<blockquote style="border-left: 4px solid ${blockquoteBorder}; padding: 0 1em; color: ${blockquoteColor}; margin: 1em 0;">`
    },
    {
      pattern: /<table>/g,
      replacement: '<table style="border-collapse: collapse; width: 100%; margin: 1em 0; border-spacing: 0;">'
    },
    {
      pattern: /<th>/g,
      replacement: `<th style="border: 1px solid ${tableBorder}; padding: 6px 13px; font-weight: bold; background-color: ${tableHeaderBackground};">`
    },
    {
      pattern: /<td>/g,
      replacement: `<td style="border: 1px solid ${tableBorder}; padding: 6px 13px;">`
    },
    {
      pattern: /<hr>/g,
      replacement: `<hr style="height: 0.25em; padding: 0; margin: 24px 0; background-color: ${hrColor}; border: 0;">`
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
    },
    {
      pattern: /<div class="markdown-it-toc">/g,
      replacement: `<div class="markdown-it-toc" style="background-color: ${tocBackground}; border: 1px solid ${tocBorder}; padding: 16px; margin: 20px 0; border-radius: 6px;">`
    },
    {
      pattern: /<ul class="markdown-it-toc-list">/g,
      replacement: '<ul class="markdown-it-toc-list" style="padding-left: 0; margin: 10px 0;">'
    },
    {
      pattern: /<li class="markdown-it-toc-item">/g,
      replacement: '<li class="markdown-it-toc-item" style="margin: 8px 0;">'
    },
    {
      pattern: /<a class="markdown-it-toc-link"/g,
      replacement: '<a class="markdown-it-toc-link" style="color: #0366d6; text-decoration: none; font-weight: 500;">'
    },
    {
      pattern: /<code>/g,
      replacement: `<code style="background-color: ${isDarkTheme ? 'rgba(255, 255, 255, 0.1)' : 'rgba(27, 31, 35, 0.05)'}; padding: 0.2em 0.4em; margin: 0; font-size: 85%; border-radius: 3px; font-family: Consolas, Monaco, \'Courier New\', monospace;">`
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
