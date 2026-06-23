using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfGridlinesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data so that gridlines are visible
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(20);

            // Ensure gridlines are shown on the sheet
            sheet.IsGridlinesVisible = true;
            // Enable printing of gridlines (required for PDF rendering)
            sheet.PageSetup.PrintGridlines = true;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use Hair gridline type – renders as a thin solid line
                GridlineType = GridlineType.Hair,
                // Optional: set a custom gridline color
                // GridlineColor = System.Drawing.Color.Black
            };

            // Save the workbook as PDF with the specified options
            workbook.Save("WorkbookWithSolidGridlines.pdf", pdfOptions);
        }
    }
}