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