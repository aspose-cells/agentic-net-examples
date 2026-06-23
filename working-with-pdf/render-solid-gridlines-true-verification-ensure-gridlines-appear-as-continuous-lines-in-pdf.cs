using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

class SolidGridlinesPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data so gridlines are visible
        worksheet.Cells["A1"].PutValue("Solid Gridlines Demo");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["C3"].PutValue(20);
        worksheet.Cells["D4"].PutValue(30);

        // Enable gridlines on the worksheet
        worksheet.IsGridlinesVisible = true;

        // Configure PDF save options to render gridlines as solid (hair) lines
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            GridlineType = GridlineType.Hair,   // hair line appears as a continuous solid line
            GridlineColor = Color.Black         // optional: ensure the gridlines are black
        };

        // Save the workbook as PDF with the specified gridline settings
        workbook.Save("SolidGridlines.pdf", pdfOptions);
    }
}