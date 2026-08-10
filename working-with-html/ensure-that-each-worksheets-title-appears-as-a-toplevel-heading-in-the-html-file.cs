// Title: Aspose.Cells for .NET – Export All Worksheets to One HTML File with Sheet Names as H1 Headings
// Description: Demonstrates how to create a workbook with several sheets, set each sheet's page header to its name (using "&A"), and save the workbook as a single HTML document. HtmlSaveOptions are configured with SaveAsSingleFile, ShowAllSheets, and ExportPageHeaders so every worksheet appears under a top‑level heading in the output HTML.
// Keywords: Aspose.Cells HTML export C# | save workbook as single HTML file | show worksheet names in HTML | ExportPageHeaders Aspose.Cells | multiple sheets to HTML | C# Aspose.Cells page header | HTMLSaveOptions ShowAllSheets
// Common Searches: Aspose.Cells export all sheets to one HTML file | How to add sheet name as heading in HTML export | C# save workbook with page headers in HTML | Show each worksheet title as H1 in Aspose.Cells HTML output | ExportPageHeaders option usage
// Developer Intent: Include every worksheet’s name as a top‑level heading in the generated HTML document.
// Use Cases: Publish a consolidated web report where each sheet’s data is introduced by a clear H1 title. | Create a single HTML manual that groups data by worksheet sections for easy navigation. | Distribute spreadsheet content via email or intranet with visible sheet titles for better readability.
// AI Prompts: Generate C# code that uses Aspose.Cells to export multiple worksheets to one HTML file with each sheet name rendered as an H1 heading. | Explain the interaction of SaveAsSingleFile, ShowAllSheets, and ExportPageHeaders in HtmlSaveOptions for rendering sheet titles. | Provide examples of customizing the font style and color of the sheet‑name headings in the exported HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook with several sheets, set each sheet's page header to its name (using "&A"), and save the workbook as a single HTML document. HtmlSaveOptions are configured with SaveAsSingleFile, ShowAllSheets, and ExportPageHeaders so every worksheet appears under a top‑level heading in the output HTML.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // Prepare worksheets with titles (names) and sample data
            // -----------------------------------------------------------------
            // First worksheet (default)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "SalesReport";
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(150);

            // Second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Stock");
            sheet2.Cells["A2"].PutValue("Banana");
            sheet2.Cells["B2"].PutValue(300);

            // Third worksheet
            Worksheet sheet3 = workbook.Worksheets.Add("Summary");
            sheet3.Cells["A1"].PutValue("Total Sales");
            sheet3.Cells["B1"].PutValue(4500);

            // -----------------------------------------------------------------
            // Configure each worksheet to display its name as a page header.
            // The header script "&A" inserts the sheet name.
            // -----------------------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Set the center section of the header to the sheet name
                ws.PageSetup.SetHeader(1, "&A");
            }

            // -----------------------------------------------------------------
            // Set HTML save options:
            //   - SaveAsSingleFile = true  : all sheets in one HTML file
            //   - ShowAllSheets = true     : render every worksheet
            //   - ExportPageHeaders = true : include the page header (sheet name)
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                SaveAsSingleFile = true,
                ShowAllSheets = true,
                ExportPageHeaders = true
            };

            // Save the workbook as HTML. Each worksheet title will appear as a
            // top‑level heading (page header) in the generated HTML.
            string outputPath = "WorkbookWithSheetHeadings.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
