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