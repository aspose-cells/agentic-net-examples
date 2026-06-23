using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data so that gridlines are visible
            worksheet.Cells["A1"].PutValue("Solid Gridlines Demo");
            worksheet.Cells["B2"].PutValue(123);
            worksheet.Cells["C3"].PutValue(DateTime.Now);

            // Ensure that gridlines are shown in the worksheet (PDF respects this setting)
            worksheet.IsGridlinesVisible = true;

            // Configure PDF save options (no RenderGridlines property needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output file path and ensure the directory exists
            string outputPath = "SolidGridlines.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}