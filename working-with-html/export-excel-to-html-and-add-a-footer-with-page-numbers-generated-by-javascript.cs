// Title: Export Excel to a Single‑File HTML page with a JavaScript‑generated page‑number footer using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill cells, set a left‑section footer containing the "Page &P" placeholder, enable footer export in HtmlSaveOptions, and save the workbook as one HTML file. The generated HTML includes JavaScript that renders the current page number in the footer.
// Keywords: Aspose.Cells | C# | .NET | HTML export | single file HTML | footer page number | ExportPageFooters | JavaScript footer | Excel to HTML
// Common Searches: Aspose.Cells export Excel to HTML with footer | Add page numbers to HTML output from Aspose.Cells .NET | Save workbook as single HTML file including footers | C# generate HTML report with Excel footer | JavaScript page number footer Aspose.Cells
// Developer Intent: Create an HTML representation of an Excel worksheet that preserves the worksheet footer and shows the current page number via JavaScript.
// Use Cases: Publish an Excel‑based report on a website while keeping printable page numbers. | Send a self‑contained HTML version of a spreadsheet via email with visible footers. | Generate a web‑ready document for archiving that retains Excel footer information.
// AI Prompts: Show how to include total pages (e.g., "Page &P of &N") in the HTML footer using Aspose.Cells. | Provide C# code to export each worksheet to its own HTML file while preserving individual footers. | Explain how to customize the CSS of the generated HTML footer for branding purposes.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, fill cells, set a left‑section footer containing the "Page &P" placeholder, enable footer export in HtmlSaveOptions, and save the workbook as one HTML file. The generated HTML includes JavaScript that renders the current page number in the footer.
class ExportExcelToHtmlWithFooter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Jane");
        worksheet.Cells["B3"].PutValue(28);

        // Set the left section of the footer to display the page number
        // &P is the placeholder for the current page number
        worksheet.PageSetup.SetFooter(0, "Page &P");

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportPageFooters = true;   // include footers in the HTML output
        saveOptions.ExportPageHeaders = true;   // optional: include headers as well
        saveOptions.SaveAsSingleFile = true;    // required for footer export

        // Save the workbook as an HTML file with the configured footer
        workbook.Save("output.html", saveOptions);
    }
}
