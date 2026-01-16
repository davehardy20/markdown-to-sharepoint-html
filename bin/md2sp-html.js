#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { Command } = require('commander');
const { markdownToSharePointHTML } = require('../index');

const program = new Command();

program
  .name('md2sp-html')
  .description('Convert Markdown files to HTML optimized for SharePoint with code block highlighting')
  .version('1.2.0')
  .argument('<input>', 'Input markdown file path')
  .argument('[output]', 'Output HTML file path (optional, prints to stdout if not provided)')
  .option('-w, --wrap', 'Wrap output in basic HTML structure with <body> tag')
  .option('-t, --toc', 'Generate table of contents')
  .option('-d, --dark-theme', 'Use dark theme for SharePoint (black/dark background)')
  .option('--toc-title <title>', 'Title for table of contents (default: "Table of Contents")')
  .action((input, output, options) => {
    try {
      const inputPath = path.resolve(input);

      if (!fs.existsSync(inputPath)) {
        console.error(`Error: Input file not found: ${inputPath}`);
        process.exit(1);
      }

      const markdown = fs.readFileSync(inputPath, 'utf-8');
      const html = markdownToSharePointHTML(markdown, {
        includeTOC: options.toc || false,
        tocTitle: options.tocTitle || 'Table of Contents',
        theme: options.darkTheme ? 'dark' : 'light'
      });

      if (output) {
        const outputPath = path.resolve(output);
        let finalHTML = html;

        if (options.wrap) {
          finalHTML = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>${path.basename(input, '.md')}</title>
</head>
<body>
${html}
</body>
</html>`;
        }

        fs.writeFileSync(outputPath, finalHTML, 'utf-8');
        console.log(`Successfully converted ${input} to ${output}`);
      } else {
        console.log(html);
      }
    } catch (error) {
      console.error(`Error: ${error.message}`);
      process.exit(1);
    }
  });

program.parse(process.argv);
