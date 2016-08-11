/**
 * Convertor options
 */
OneNoteToMarkdown.options = {
    
    /**
     * Debug logging
     */
    "debug": false,

    /**
     * If true, title will be added as a H1 and heading 1 as H2 and so on
     * If false title will be added as a paragraph and headings remain the same
     */
    "treatTitleAsH1": true,

    /**
     * If true, headings will be closed with hashes (### heading **###** <--- this!)
     */
    "closeHeadings": false,

    /**
     * Character to be used as the <ul> bullet
     */
    "listBullets": "*",

    /**
     * Indentation character(s)
     */
    "indentation": "\t",

    /**
     * What type should be copied from clipboard
     *
     * One note copies content in HTML, however if you have the HTML
     * in a text editor (say you have it from the .mht file) you can
     * set this property to "text/plain"
     */
    "clipboardType": "text/html",

    /**
     * Force parser to use a given OneNote version
     *
     * Useful if the HTML is not copied from OneNote or it's
     * an older OneNote document openned in a newer application
     */
    "forcedVersion": null,

    /**
     * Regexes to match heading types from html style attribute
     * on the <p> tag for different OneNote versions
     *
     * - h0 is title
     * - order them by regex specificity to avoid bad matching (more general selectors put on bottom)
     */
    "oneNoteVersions": {

        // MS Office 2013
        "Microsoft OneNote 15": [
            ["font-family:\\s*\"Calibri Light\"", "h0"],
            ["font-size:\\s*12.*color:\\s*#5B9BD5.*font-style", "h4"],
            ["font-size:\\s*11.*color:\\s*#2E75B5.*font-style", "h6"],
            ["font-size:\\s*16.*color:\\s*#1E4E79", "h1"],
            ["font-size:\\s*14.*color:\\s*#2E75B5", "h2"],
            ["font-size:\\s*12.*color:\\s*#5B9BD5", "h3"],
            ["font-size:\\s*11.*color:\\s*#2E75B5", "h5"],
        ],

        // MS Office 2010
        "Microsoft OneNote 14": [
            ["font-size:\\s*17", "h0"],
            ["font-size:\\s*16", "h1"],
            ["font-size:\\s*13", "h2"],
            ["font-size:\\s*11.*color.*font-weight.*font-style", "h4"],
            ["font-size:\\s*11.*color.*font-weight", "h3"],
            ["font-size:\\s*11.*color.*font-style", "h6"],
            ["font-size:\\s*11.*color", "h5"],
        ]
    }
}

//////////////////////
// Helper functions //
//////////////////////

/**
 * Checks wether a given OneNote version is supported
 */
OneNoteToMarkdown.options.supportsOneNoteVersion = function(version)
{
    if (OneNoteToMarkdown.options.oneNoteVersions[version] === undefined)
        throw new Error("Unsupported OneNote version: " + version)
}