/*
 * OneNoteToMarkdown.js 0.8.0
 *
 * https://github.com/Jirka-Mayer/OneNoteToMarkdown
 *
 * Copyright (c) 2016 Jiri Mayer
 * Licensed under the MIT license.
 */

+function()
{

var OneNoteToMarkdown = {};

/**
 * Converts the OneNote html output to markdown
 */
OneNoteToMarkdown.convert = function(html)
{
    // parse HTML into DOM
    var div = document.createElement("div")
    div.innerHTML = html

    // version of the OneNote
    var oneNoteVersion = OneNoteToMarkdown.options.forcedVersion
    var versionMatches = null

    if (oneNoteVersion !== null)
    {
        OneNoteToMarkdown.options.supportsOneNoteVersion(oneNoteVersion)
        versionMatches = OneNoteToMarkdown.options.oneNoteVersions[oneNoteVersion]
    }

    // mardown builder
    var builder = new OneNoteToMarkdown.MDBuilder()

    // go through all elements
    for (var i in div.children)
    {
        var child = div.children[i]

        // handle meta tags
        if (child.tagName == "META")
        {
            // get OneNote version and check that we support it
            if (child.name == "Generator")
            {
                // ignore if already set
                if (oneNoteVersion !== null)
                    continue

                oneNoteVersion = child.content
                OneNoteToMarkdown.options.supportsOneNoteVersion(oneNoteVersion)
                versionMatches = OneNoteToMarkdown.options.oneNoteVersions[oneNoteVersion]
            }

            continue
        }
        else if (oneNoteVersion === null) // no meta tag but unknown generator - error
            throw new Error("OneNote version not provided. Are you really copying from a OneNote?")

        // handle paragraphs and headings
        if (child.tagName == "P")
        {
            OneNoteToMarkdown.convert_handleParagraphTag(child, builder, versionMatches)
            continue
        }

        // handle lists
        if (child.tagName == "OL" || child.tagName == "UL")
        {
            OneNoteToMarkdown.convert_handleListTag(child, builder)
            continue
        }
    }

    if (OneNoteToMarkdown.options.debug)
    {
        console.log("Input HTML (accessible through 'window.div'):")
        console.log(div)
        window.div = div
    }

    return builder.output
}

OneNoteToMarkdown.convert_handleParagraphTag = function(elem, builder, versionMatches)
{
    var style = elem.getAttribute("style")

    if (style === null) // no style attrib => treat as paragraph
    {
        builder.addParagraph(elem.innerText)
        console.warn("Unknown paragraph tag, treating as a paragraph.")
    }

    // get child style (span) and append it
    if (elem.children.length == 1 && elem.children[0] && elem.children[0].tagName == "SPAN")
    {
        if (elem.children[0].style)
            style += ";" + elem.children[0].getAttribute("style") || ""
    }

    // find a type match
    var type = "p"
    for (var i in versionMatches)
    {
        if (style.match(versionMatches[i][0]))
        {
            type = versionMatches[i][1]
            break
        }
    }

    // add a paragraph
    if (type == "p")
        builder.addParagraph(elem.innerHTML)
    else // add a heading
        builder.addHeading(type, elem.innerText)
}

OneNoteToMarkdown.convert_handleListTag = function(elem, builder, depth)
{
    if (depth === undefined)
        depth = 0

    if (depth == 0)
        builder.topMargin(1)

    // is this list ordered?
    var ordered = elem.tagName == "OL"

    // <ol> index
    var index = 0

    for (var i in elem.children)
    {
        var child = elem.children[i]

        if (child.tagName == "LI")
        {
            index++

            if (ordered)
                builder.addOL(index, depth, child.innerHTML)
            else
                builder.addUL(depth, child.innerHTML)

            continue
        }

        if (child.tagName == "OL" || child.tagName == "UL")
        {
            OneNoteToMarkdown.convert_handleListTag(child, builder, depth + 1)
            continue
        }
    }

    if (depth == 0)
        builder.bottomMargin(1)
}
/**
 * Built-in function for handling paste event
 */
OneNoteToMarkdown.handlePaste = function(e)
{
    // type to get from the clipboard
    var targetType = OneNoteToMarkdown.options.clipboardType
    
    // check chlipboard API
    if (!e || !e.clipboardData || !e.clipboardData.types || !e.clipboardData.getData)
        throw Error("Your browser seems to be too old.")

    // all types in clipboard
    var types = e.clipboardData.types;

    // check that we have required type in the clipboard
    if (((types instanceof DOMStringList) && !types.contains(targetType)) || (types.indexOf && types.indexOf(targetType) === -1))
        throw new Error("There's no '" + targetType + "' in the clipboard, only: " + JSON.stringify(types))
    
    // get the html code
    var html = e.clipboardData.getData(targetType)

    // convert to markdown
    var md = OneNoteToMarkdown.convert(html)

    return md
}
OneNoteToMarkdown.MDBuilder = function()
{
    /**
     * Output markdown
     */
    this.output = ""

    /**
     * How many empty lines are there as the last margin
     */
    this.lastMargin = Number.POSITIVE_INFINITY // make sure there's no unnecessary top margin
}

/**
 * Adds a paragraph to the output
 */
OneNoteToMarkdown.MDBuilder.prototype.addParagraph = function(content)
{
    content = this.parseEmphasis(content)
    content = this.cureContent(content)

    this.topMargin(1)

    this.output += "\n" + content

    this.bottomMargin(1)
}

/**
 * Adds a heading or title to the output
 */
OneNoteToMarkdown.MDBuilder.prototype.addHeading = function(type, content)
{
    content = this.cureContent(content)

    // get hading number
    var headingNumber = parseInt(type[1])

    if (OneNoteToMarkdown.options.treatTitleAsH1)
    {
        // increment heading numbers
        headingNumber++

        // h7 is too high, add paragraph instead
        if (headingNumber == 7)
        {
            this.addParagraph(content)
            return
        }
    }
    else
    {
        // treat the special snowflake as a paragraph
        if (type == "h0")
        {
            this.addParagraph(content)
            return
        }
    }

    this.topMargin(2)

    this.output += "\n"

    for (var i = 0; i < headingNumber; i++)
        this.output += "#"

    this.output += " " + content

    if (OneNoteToMarkdown.options.closeHeadings)
    {
        this.output += " "

        for (var i = 0; i < headingNumber; i++)
            this.output += "#"
    }

    this.bottomMargin(1)
}

/**
 * Adds an ordered list bullet to the output
 */
OneNoteToMarkdown.MDBuilder.prototype.addOL = function(index, depth, content)
{
    content = this.parseEmphasis(content)
    content = this.cureContent(content)

    this.output += "\n"

    for (var i = 0; i < depth; i++)
        this.output += OneNoteToMarkdown.options.indentation
    
    this.output += index + ". "
    this.output += content
}

/**
 * Adds an unordered list bullet to the output
 */
OneNoteToMarkdown.MDBuilder.prototype.addUL = function(depth, content)
{
    content = this.parseEmphasis(content)
    content = this.cureContent(content)

    this.output += "\n"

    for (var i = 0; i < depth; i++)
        this.output += OneNoteToMarkdown.options.indentation

    this.output += OneNoteToMarkdown.options.listBullets + " "
    this.output += content
}

/**
 * Does some whitespace fixing
 */
OneNoteToMarkdown.MDBuilder.prototype.cureContent = function(text)
{
    return text
        .replace(/\s+/g, " ") // collapse whitespace to a single space
        .replace(/\s*<br>\s*/, "<br>") // remove whitespace around <br> tags
}

/**
 * Replaces "bold" and "italic" - like tags
 * (OneNote exports them as <span> to make my life easier)
 */
OneNoteToMarkdown.MDBuilder.prototype.parseEmphasis = function(html)
{
    return html
        .replace(/<span[^>]*?bold[^>]*?italic[^>]*?>([^<]*?)<\/span>/g, "***$1***")
        .replace(/<span[^>]*?italic[^>]*?>([^<]*?)<\/span>/g, "*$1*")
        .replace(/<span[^>]*?bold[^>]*?>([^<]*?)<\/span>/g, "**$1**")

        // remove other spans
        .replace(/<span[^>]*>([\s\S]*?)<\/span>/g, "$1")

        // replace actual braces that were escaped
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
}

/**
 * Makes sure there's a given margin on the top of an element that's gonna be added
 */
OneNoteToMarkdown.MDBuilder.prototype.topMargin = function(size)
{
    // how many lines to add
    var add = size - this.lastMargin
    
    for (var i = 0; i < add; i++)
        this.output += "\n"
}

/**
 * Makes sure there's a given margin on the bottom of the last added element
 */
OneNoteToMarkdown.MDBuilder.prototype.bottomMargin = function(size)
{
    for (var i = 0; i < size; i++)
        this.output += "\n"

    this.lastMargin = size
}
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

////////////////////////////
// Register to the window //
////////////////////////////

window.OneNoteToMarkdown = OneNoteToMarkdown;

}();