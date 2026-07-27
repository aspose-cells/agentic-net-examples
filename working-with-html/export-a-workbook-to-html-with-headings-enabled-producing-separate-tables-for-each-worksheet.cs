// Title: C# Aspose.Cells: Export Workbook to HTML with Row/Column Headings and Separate Files
// Description: Demonstrates how to use Aspose.Cells HtmlSaveOptions in C# to enable ExportRowColumnHeadings, disable SaveAsSingleFile, and generate individual HTML files for each worksheet, preserving column letters and row numbers.
// Keywords: Aspose.Cells HTML export | C# HtmlSaveOptions | ExportRowColumnHeadings | SaveAsSingleFile false | multiple worksheets to HTML | separate HTML files per sheet | HTML column headings Aspose | Excel to HTML C# example
// Common Searches: Aspose.Cells export workbook to HTML with headings C# | Save each worksheet as separate HTML file Aspose.Cells | HtmlSaveOptions ExportRowColumnHeadings example | How to disable SaveAsSingleFile in Aspose.Cells | C# convert Excel sheets to individual HTML tables
// Developer Intent: Generate HTML output that includes spreadsheet row/column headings and creates one HTML file for each worksheet.
// Use Cases: Create printable HTML reports that retain Excel column letters and row numbers. | Build web documentation where each sheet is loaded as an independent HTML table. | Serve separate HTML files per sheet for dynamic loading in single‑page applications.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with ExportRowColumnHeadings enabled and each worksheet saved as a separate file. | Explain the impact of HtmlSaveOptions.SaveAsSingleFile = false when exporting multiple worksheets to HTML. | Show how to specify a custom output folder and naming pattern for each worksheet's HTML file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to use Aspose.Cells HtmlSaveOptions in C# to enable ExportRowColumnHeadings, disable SaveAsSingleFile, and generate individual HTML files for each worksheet, preserving column letters and row numbers.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "FirstSheet";
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");

            // Populate first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Cells["A1"].PutValue("Header1");
            sheet1.Cells["B1"].PutValue("Header2");
            sheet1.Cells["A2"].PutValue("Data1");
            sheet1.Cells["B2"].PutValue("Data2");

            // Populate second worksheet
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Quantity");
            sheet2.Cells["A2"].PutValue("Apple");
            sheet2.Cells["B2"].PutValue(10);
            sheet2.Cells["A3"].PutValue("Banana");
            sheet2.Cells["B3"].PutValue(20);

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                // Export row and column headings (A, B, 1, 2, …)
                ExportRowColumnHeadings = true,
                // Ensure each worksheet is saved as a separate HTML file
                SaveAsSingleFile = false
            };

            // Save the workbook to HTML. Separate files will be generated for each sheet.
            workbook.Save("WorkbookWithHeadings.html", options);
        }
    }
}
