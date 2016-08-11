OneNote to markdown
===================

OneNoteToMarkdown is a javascript library for converting OneNote documents into markdown.

Example of usage can be found in [examples/simple.html](./examples/simple.html)

But if you just want to convert some documents, this library is used [here](http://jirka-mayer.github.io/onenote-to-markdown)

## What works

Library supports copying content from OneNote 2010 and 2013, however you can add custom parsing conditions. Just take a look at [options file](./src/options.js).

What works:

- Paragraphs
    - Bold
    - Italic
- Title
- Heading 1 to 6
- `ul` and `ol`

## What doesn't work

I've built this tool for myself for one-time use, so it only supports features I needed. But feel free to explore and extend to match your needs. It's not a lot of code and PRs are welcomed.

What doesn't work:

- Tables
- Code
- Quotes
- Images
    - no idea what would happen if you tried
- Links
    - well I didn't test that but I think they should go through as a plain HTML code
- Other OneNote specific fancy stuff