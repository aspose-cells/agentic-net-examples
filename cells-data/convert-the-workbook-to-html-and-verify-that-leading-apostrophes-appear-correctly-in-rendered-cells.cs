// Title: Export Aspose.Cells Workbook to HTML with Correct Leading Apostrophe Rendering (C#)
// Description: Demonstrates how to enable QuotePrefix handling, write a value that begins with a single quote, save the workbook as HTML, and verify that the leading apostrophe appears correctly (as &#39; or a literal character) in the generated HTML. Includes code to read the HTML file and retrieve the cell's HTML string.
// Keywords: Aspose.Cells | C# | HTML export | leading apostrophe | QuotePrefixToStyle | QuotePrefix style | cell HTML string | preserve apostrophe | Aspose.Cells HTML save
// Common Searches: Aspose.Cells preserve leading apostrophe in HTML | QuotePrefixToStyle HTML export Aspose.Cells | how to show single quote in exported HTML spreadsheet | Aspose.Cells GetHtmlString apostrophe | verify apostrophe entity in Aspose.Cells HTML output
// Developer Intent: Generate HTML from a workbook and confirm that cells starting with an apostrophe retain the apostrophe in the exported HTML.
// Use Cases: Export spreadsheets to web‑ready HTML while keeping leading apostrophes visible for data that uses them as text qualifiers. | Automated testing to ensure the exported HTML contains the correct apostrophe entity, guaranteeing data integrity after conversion. | Extract a cell's HTML representation for embedding in emails, reports, or custom web components.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as HTML and check that a leading apostrophe in cell A1 is rendered as &#39; or a literal quote. | Explain the effect of Workbook.Settings.QuotePrefixToStyle and Style.QuotePrefix on HTML export of cells with leading apostrophes. | Create a C# unit test that creates a workbook with a leading apostrophe, exports it to HTML, and asserts the presence of the apostrophe entity in the output.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to enable QuotePrefix handling, write a value that begins with a single quote, save the workbook as HTML, and verify that the leading apostrophe appears correctly (as &#39; or a literal character) in the generated HTML. Includes code to read the HTML file and retrieve the cell's HTML string.
class LeadingApostropheHtmlDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable automatic QuotePrefix handling for values that start with a single quote
        workbook.Settings.QuotePrefixToStyle = true;

        // Set a cell value that begins with a leading apostrophe
        worksheet.Cells["A1"].PutValue("'Hello");

        // Ensure the style reflects the QuotePrefix (optional, shown for clarity)
        Style style = worksheet.Cells["A1"].GetStyle();
        style.QuotePrefix = true;
        worksheet.Cells["A1"].SetStyle(style);

        // Save the workbook as HTML using default HtmlSaveOptions
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        string htmlPath = "output.html";
        workbook.Save(htmlPath, htmlOptions);

        // Read the generated HTML file
        string htmlContent = File.ReadAllText(htmlPath);

        // Verify that the leading apostrophe is present in the HTML output
        // It may appear as an HTML entity (&#39;) or as a literal character
        bool apostrophePresent = htmlContent.Contains("&#39;") || htmlContent.Contains("'Hello");
        Console.WriteLine("HTML contains leading apostrophe: " + apostrophePresent);

        // Optionally, display the cell's own HTML representation
        string cellHtml = worksheet.Cells["A1"].GetHtmlString(true);
        Console.WriteLine("Cell HTML string: " + cellHtml);
    }
}
