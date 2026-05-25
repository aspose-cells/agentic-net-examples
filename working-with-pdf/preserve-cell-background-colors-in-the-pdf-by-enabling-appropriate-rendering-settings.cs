using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class PreserveCellBackground
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a solid background style for cell A1
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;          // solid fill
            style.ForegroundColor = Color.Yellow;          // background color (foreground used for solid pattern)

            // Apply style to cell A1
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Cell with yellow background");
            cell.SetStyle(style);

            // Configure PDF save options (cell shading is preserved by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output file path
            string outputPath = "CellBackgroundPreserved.pdf";

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}