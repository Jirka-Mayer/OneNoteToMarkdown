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