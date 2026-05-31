using System;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data so that gridlines are visible
        worksheet.Cells["A1"].PutValue("Solid Gridlines Demo");
        worksheet.Cells["B2"].PutValue(123);
        worksheet.Cells["C3"].PutValue(456);

        // Enable gridlines visibility on the worksheet
        worksheet.IsGridlinesVisible = true;

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set the gridline type to a solid line (Hair)
        pdfSaveOptions.GridlineType = GridlineType.Hair;

        // Optionally set the gridline color (default is black)
        pdfSaveOptions.GridlineColor = Color.Black;

        // Save the workbook as a PDF with solid gridlines
        workbook.Save("SolidGridlines.pdf", pdfSaveOptions);
    }
}