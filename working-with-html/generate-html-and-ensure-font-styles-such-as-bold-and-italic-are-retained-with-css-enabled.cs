// Title: Export Excel to HTML with CSS while preserving <b> and <i> formatting using Aspose.Cells for .NET
// Description: This example creates a workbook, inserts HTML‑styled text (<b>Bold</b> and <i>Italic</i>) into cell A1, configures HtmlSaveOptions (ParseHtmlTagInCell, DisableCss = false, AddGenericFont = true) and saves the file as StyledOutput.html. The resulting HTML uses CSS to render bold and italic text exactly as defined in the cell.
// Keywords: Aspose.Cells HTML export .NET | ParseHtmlTagInCell | CSS styling Excel to HTML | preserve bold italic Aspose | HtmlSaveOptions example | C# Excel to HTML conversion | generic font fallback CSS
// Common Searches: Aspose.Cells keep <b> tags when saving as HTML | How to enable CSS for Excel to HTML export .NET | Parse HTML tags inside Excel cells with Aspose | Export workbook to HTML with bold and italic styling | C# Aspose.Cells HtmlSaveOptions settings
// Developer Intent: Generate an HTML file from an Excel workbook that retains bold and italic styling defined by HTML tags inside cell values, using CSS via Aspose.Cells for .NET.
// Use Cases: Publishing web‑ready reports where cell content includes HTML markup and the exported page must display bold/italic text correctly. | Automating conversion of Excel templates with embedded HTML formatting into browser‑compatible HTML with fallback fonts. | Building a .NET service that returns styled HTML snippets for email or web templates while preserving inline HTML tag formatting.
// AI Prompts: Show how to modify HtmlSaveOptions to embed all CSS inline while still parsing HTML tags in cells. | Give an example of adding custom CSS classes to the generated HTML for bold and italic text using Aspose.Cells. | Explain how to export each worksheet to a separate HTML file with the same CSS and tag‑parsing configuration.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // This example creates a workbook, inserts HTML‑styled text (<b>Bold</b> and <i>Italic</i>) into cell A1, configures HtmlSaveOptions (ParseHtmlTagInCell, DisableCss = false, AddGenericFont = true) and saves the file as StyledOutput.html. The resulting HTML uses CSS to render bold and italic text exactly as defined in the cell.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put HTML formatted text into a cell.
            // The HTML tags <b> and <i> will be parsed and rendered as bold and italic.
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("<b>Bold Text</b> and <i>Italic Text</i>");

            // Configure HTML save options.
            // DisableCss = false ensures that CSS styles are used (default behavior).
            // ParseHtmlTagInCell = true makes the HTML tags inside the cell value be interpreted.
            // AddGenericFont = true (default) adds a generic font fallback in the CSS.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                DisableCss = false,
                ParseHtmlTagInCell = true,
                AddGenericFont = true
            };

            // Save the workbook as an HTML file with the specified options.
            workbook.Save("StyledOutput.html", htmlOptions);
        }
    }
}
