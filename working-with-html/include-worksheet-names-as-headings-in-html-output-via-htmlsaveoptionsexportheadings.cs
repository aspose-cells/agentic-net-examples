// Title: Export Multiple Worksheets to a Single HTML File with Sheet Tabs Using Aspose.Cells for .NET
// Description: This example builds a workbook with two sheets (SalesData and Inventory), fills them with sample data, and uses HtmlSaveOptions (ExportRowColumnHeadings, SaveAsSingleFile, ShowAllSheets) to generate one HTML page where each worksheet is displayed as a tab heading and row/column headings (A, B, 1, 2…) are preserved.
// Keywords: Aspose.Cells HTML export C# | save workbook as single HTML | sheet tabs in HTML | ExportRowColumnHeadings .NET | ShowAllSheets Aspose.Cells | SaveAsSingleFile HTML | Excel to HTML with headings | C# .NET spreadsheet web view
// Common Searches: Aspose.Cells export multiple sheets to one HTML file | C# HtmlSaveOptions ShowAllSheets example | How to add worksheet tabs when saving Excel as HTML | ExportRowColumnHeadings vs ExportHeadings Aspose.Cells | Create HTML report with sheet names as headings C#
// Developer Intent: Generate a single HTML document that presents each worksheet as a clickable tab with row and column headings.
// Use Cases: Publish a sales and inventory dashboard on a website without requiring Excel. | Provide a read‑only, web‑friendly view of multi‑sheet reports for global teams. | Embed Excel‑style navigation in online documentation or intranet portals.
// AI Prompts: Show how to style the HTML tab bar for worksheet headings using Aspose.Cells HtmlSaveOptions. | Explain when to use ExportRowColumnHeadings instead of the deprecated ExportHeadings property. | Create C# code that saves each worksheet to a separate HTML file while keeping the sheet name as a heading.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // This example builds a workbook with two sheets (SalesData and Inventory), fills them with sample data, and uses HtmlSaveOptions (ExportRowColumnHeadings, SaveAsSingleFile, ShowAllSheets) to generate one HTML page where each worksheet is displayed as a tab heading and row/column headings (A, B, 1, 2…) are preserved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Rename the default worksheet and add sample data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "SalesData";
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(150);
            sheet1.Cells["A3"].PutValue("Banana");
            sheet1.Cells["B3"].PutValue(200);

            // Add a second worksheet with its own data
            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Stock");
            sheet2.Cells["A2"].PutValue("Pen");
            sheet2.Cells["B2"].PutValue(500);
            sheet2.Cells["A3"].PutValue("Notebook");
            sheet2.Cells["B3"].PutValue(300);

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                // Export row and column headings (A, B, 1, 2, …)
                ExportHeadings = true,                 // Obsolete but still functional
                ExportRowColumnHeadings = true,        // Preferred property
                // Save all worksheets into a single HTML file so that each sheet name appears as a tab heading
                SaveAsSingleFile = true,
                ShowAllSheets = true
            };

            // Save the workbook as HTML; each worksheet name will appear as a tab heading
            workbook.Save("WorkbookWithSheetHeadings.html", options);
        }
    }
}
