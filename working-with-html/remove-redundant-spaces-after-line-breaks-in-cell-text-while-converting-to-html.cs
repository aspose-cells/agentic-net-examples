// Title: C# – Remove Redundant Spaces After <br> Tags When Converting Excel Cells to HTML with Aspose.Cells
// Description: Demonstrates how to export a worksheet containing multiline text to HTML using Aspose.Cells, enable text wrapping so line breaks become <br> tags, and then clean up any whitespace that follows those tags with a regular expression.
// Keywords: Aspose.Cells HTML export | C# remove spaces after br | regex clean HTML whitespace | text wrapping Excel to HTML | line break spacing issue
// Common Searches: remove extra spaces after <br> in Aspose.Cells HTML output | Aspose.Cells C# line break whitespace cleanup | HTMLSaveOptions remove whitespace after line breaks | regex to trim spaces after br tag in generated HTML | export Excel to HTML without redundant spaces
// Developer Intent: Eliminate unnecessary whitespace that appears after <br> tags in HTML generated from an Excel workbook using Aspose.Cells.
// Use Cases: Export wrapped text with line breaks to HTML and ensure clean markup. | Post‑process Aspose.Cells HTML output with a regex to improve layout consistency. | Automate report generation where HTML files must not contain stray spaces after line‑break tags.
// AI Prompts: Show C# code that uses Aspose.Cells to export a worksheet to HTML and then removes spaces after <br> tags with a regular expression. | Explain how to enable text wrapping in Aspose.Cells so line breaks are rendered as <br> elements in the HTML file. | Suggest alternative approaches to trim whitespace after line breaks without reading and rewriting the HTML file.

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Demonstrates how to export a worksheet containing multiline text to HTML using Aspose.Cells, enable text wrapping so line breaks become <br> tags, and then clean up any whitespace that follows those tags with a regular expression.
class RemoveRedundantSpacesDemo
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put text containing line breaks (\n) and redundant spaces after them
        worksheet.Cells["A1"].PutValue("First line\n   Second line\n    Third line");

        // Enable text wrapping so that line breaks are exported as <br> tags in HTML
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(wrapStyle);

        // Save the workbook to HTML
        string htmlPath = "output.html";
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        workbook.Save(htmlPath, saveOptions);

        // Load the generated HTML, remove spaces that follow <br> tags, and overwrite the file
        string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);
        string cleanedHtml = Regex.Replace(htmlContent, @"<br>\s+", "<br>");
        File.WriteAllText(htmlPath, cleanedHtml, Encoding.UTF8);

        Console.WriteLine("HTML file saved with redundant spaces after line breaks removed.");
    }
}
