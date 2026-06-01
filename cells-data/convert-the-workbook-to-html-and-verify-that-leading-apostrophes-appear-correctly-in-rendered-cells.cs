using System;
using System.IO;
using Aspose.Cells;

class LeadingApostropheHtmlDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable automatic QuotePrefix handling so leading apostrophes are treated as style
        workbook.Settings.QuotePrefixToStyle = true;

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a value that starts with a single quote
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("'Hello");

        // Save the workbook as HTML
        string htmlPath = "output.html";
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(); // default options
        workbook.Save(htmlPath, htmlOptions);

        // Load the generated HTML file
        string htmlContent = File.ReadAllText(htmlPath);

        // Verify that the leading apostrophe appears in the HTML output
        // Aspose may encode the apostrophe as &#39; or keep it as a literal character
        bool containsApostrophe = htmlContent.Contains("&#39;") || htmlContent.Contains("'Hello");
        Console.WriteLine("HTML contains leading apostrophe: " + containsApostrophe);

        // Additional verification using the cell's GetHtmlString method
        string cellHtml = cell.GetHtmlString(true);
        Console.WriteLine("Cell HTML string: " + cellHtml);
    }
}