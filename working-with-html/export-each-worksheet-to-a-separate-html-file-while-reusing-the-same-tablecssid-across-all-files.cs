// Title: Export each worksheet to its own HTML file with a shared TableCssId using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, assign a common TableCssId, configure HtmlSaveOptions (ExportActiveWorksheetOnly, ExportWorksheetCSSSeparately), loop through all original worksheets, set each as active, and save them as separate HTML files named after the worksheet while reusing the same CSS identifier.
// Keywords: Aspose.Cells HTML export | TableCssId reuse | Export worksheets to HTML .NET | HtmlSaveOptions ExportActiveWorksheetOnly | Separate HTML per sheet | C# Aspose.Cells example | consistent table styling HTML
// Common Searches: Aspose.Cells export each sheet to HTML | shared TableCssId for multiple HTML files | loop through worksheets and save as HTML C# | how to use HtmlSaveOptions for per‑sheet export | consistent CSS id across exported HTML pages
// Developer Intent: Generate individual HTML pages for every worksheet in a workbook while applying the same TableCssId to maintain uniform table styling.
// Use Cases: Publish Excel reports as separate web pages with a common table style. | Create documentation sets where each sheet becomes an HTML file sharing identical CSS identifiers. | Automate batch conversion of workbooks for a portal that requires consistent table styling across all pages.
// AI Prompts: Write C# code that iterates through all worksheets in an Aspose.Cells workbook and saves each as an HTML file using a single TableCssId. | Explain the interaction between HtmlSaveOptions.ExportActiveWorksheetOnly and TableCssId when exporting multiple worksheets. | Modify the sample to generate an external CSS file per worksheet while keeping the same TableCssId in each HTML output.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, assign a common TableCssId, configure HtmlSaveOptions (ExportActiveWorksheetOnly, ExportWorksheetCSSSeparately), loop through all original worksheets, set each as active, and save them as separate HTML files named after the worksheet while reusing the same CSS identifier.
class ExportWorksheetsToHtml
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data to three worksheets
            Workbook workbook = new Workbook();

            // First worksheet (default)
            workbook.Worksheets[0].Name = "Sheet1";
            workbook.Worksheets[0].Cells["A1"].PutValue("Data in Sheet1");

            // Second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Data in Sheet2");

            // Third worksheet
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
            sheet3.Cells["A1"].PutValue("Data in Sheet3");

            // Define a common TableCssId that will be used for all HTML files
            const string commonTableCssId = "my-table-style";

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,   // Export only the active worksheet each time
                TableCssId = commonTableCssId,      // Apply the same TableCssId to all files
                ExportWorksheetCSSSeparately = true // Export CSS for each worksheet separately
            };

            // Capture the original worksheet count to avoid processing evaluation warning sheets
            int originalSheetCount = workbook.Worksheets.Count;

            // Iterate through each original worksheet, set it as active, and save to an individual HTML file
            for (int i = 0; i < originalSheetCount; i++)
            {
                workbook.Worksheets.ActiveSheetIndex = i;
                string htmlFileName = $"{workbook.Worksheets[i].Name}.html";

                // Ensure the target path is valid (optional safety check)
                string fullPath = Path.GetFullPath(htmlFileName);
                workbook.Save(fullPath, saveOptions);
                Console.WriteLine($"Saved {htmlFileName} with TableCssId = \"{commonTableCssId}\"");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
