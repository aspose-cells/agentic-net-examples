using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class SolidGridlinesPdfDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data so gridlines are visible
            worksheet.Cells["A1"].PutValue("Solid Gridlines Demo");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue(2);
            worksheet.Cells["C2"].PutValue(3);
            worksheet.Cells["A3"].PutValue(4);
            worksheet.Cells["B3"].PutValue(5);
            worksheet.Cells["C3"].PutValue(6);

            // Ensure worksheet gridlines are turned on
            worksheet.IsGridlinesVisible = true;

            // Configure PDF save options (solid gridlines)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                GridlineType = GridlineType.Hair,   // thin solid line
                GridlineColor = Color.Black
            };

            // Define output file path
            string outputPath = "SolidGridlinesDemo.pdf";

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