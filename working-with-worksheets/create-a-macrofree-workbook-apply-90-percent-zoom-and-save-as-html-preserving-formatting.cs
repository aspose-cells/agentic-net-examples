using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new, macro‑free workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just to have content)
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(15);

            // Set the worksheet zoom level to 90%
            sheet.PageSetup.Zoom = 90;

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable scaling based on the worksheet zoom level
            htmlOptions.WorksheetScalable = true;

            // Preserve grid lines and other formatting (default values already preserve most formatting)
            htmlOptions.ExportGridLines = true;

            // Save the workbook as HTML
            workbook.Save("Workbook90PercentZoom.html", htmlOptions);
        }
    }
}