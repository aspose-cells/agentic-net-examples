// Title: Export a workbook to HTML with Aspose.Cells in C# while preserving leading apostrophes in cell values
// AI Prompts: Write C# code that creates a workbook, sets a cell value beginning with an apostrophe, enables QuotePrefix, saves the workbook as HTML using Aspose.Cells, and confirms the apostrophe appears in the output. | Show how to obtain the HTML string of a single cell after applying QuotePrefix with Aspose.Cells. | Provide a snippet that reads the generated HTML file and verifies the leading apostrophe is present as a literal character or as the &#39; entity.
// Common Searches: Aspose.Cells C# export to HTML keep leading single quote in cell | How to display a leading apostrophe in a cell when saving as HTML with Aspose.Cells | QuotePrefix property effect on HTML output in Aspose.Cells | Get HTML string for a specific cell after setting QuotePrefix in Aspose.Cells | Validate apostrophe character in generated HTML file using Aspose.Cells C#
// Tags: html export quote-prefix aspocells c# | preserve leading apostrophe aspocells html | cell gethtmlstring aspocells | verify apostrophe in html output aspocells | save workbook as html aspocells c#

using System;
using System.IO;
using Aspose.Cells;
using System.Text.RegularExpressions;

// The example creates a new workbook, writes a value that starts with an apostrophe into cell A1, applies the QuotePrefix style so the apostrophe is treated as a literal character, saves the workbook as HTML with Aspose.Cells, reads the generated HTML file to check for the apostrophe (either as a literal ' or the &#39; entity), and prints both the verification result and the HTML representation of the cell.
class ExportWorkbookToHtmlWithApostrophe
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set a cell value that starts with a leading apostrophe
        // The apostrophe is stored as part of the text and should be displayed in HTML
        Cell cell = cells["A1"];
        cell.PutValue("'LeadingApostrophe");

        // Enable QuotePrefix style so the leading apostrophe is treated as a literal character
        Style style = cell.GetStyle();
        style.QuotePrefix = true;
        cell.SetStyle(style);

        // Prepare HTML save options (default options are sufficient for this scenario)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Define output HTML file path
        string htmlPath = "WorkbookWithApostrophe.html";

        // Save the workbook as HTML
        workbook.Save(htmlPath, htmlOptions);

        // Read the generated HTML content
        string htmlContent = File.ReadAllText(htmlPath);

        // Verify that the leading apostrophe appears in the HTML output
        // It may be encoded as &#39; or appear as a literal '
        bool containsApostrophe = htmlContent.Contains("'") || htmlContent.Contains("&#39;");

        Console.WriteLine("HTML contains leading apostrophe: " + containsApostrophe);

        // Additionally, get the HTML string for the specific cell and display it
        string cellHtml = cell.GetHtmlString(true);
        Console.WriteLine("Cell A1 HTML representation:");
        Console.WriteLine(cellHtml);
    }
}
