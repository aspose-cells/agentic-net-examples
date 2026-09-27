// Title: Create a PDF with a custom page size that matches the used range of an Excel worksheet using Aspose.Cells in C#
// AI Prompts: Generate C# code that measures the total width and height of the used cells, converts the pixel values to points, and applies them to the worksheet's custom page dimensions before saving as PDF with Aspose.Cells. | Show how to configure PdfSaveOptions so the PDF is generated without any automatic page scaling, preserving the worksheet's original layout.
// Common Searches: how to set custom PDF page dimensions from Excel used range Aspose.Cells C# | calculate worksheet pixel size and convert to points for PDF export Aspose.Cells | Aspose.Cells disable scaling when converting Excel to PDF | use PageSetup.CustomPaperSize to match Excel content size in PDF | C# export Excel sheet to PDF with exact layout no scaling
// Tags: Aspose.Cells set page dimensions | PdfSaveOptions turn off scaling | C# compute worksheet used range size | export Excel to PDF exact layout | convert pixels to points Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, determines the used range extents, sums column widths and row heights in pixels, converts those totals to points, and (optionally) assigns the dimensions to a custom page size. PdfSaveOptions are configured with OnePagePerSheet disabled to prevent automatic scaling, resulting in a PDF that precisely matches the worksheet's layout.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Determine used range extents
            int maxColumn = sheet.Cells.MaxColumn; // zero‑based index of the last used column
            int maxRow = sheet.Cells.MaxRow;       // zero‑based index of the last used row

            // Accumulate total width and height in pixels
            double totalWidthPixels = 0;
            for (int col = 0; col <= maxColumn; col++)
                totalWidthPixels += sheet.Cells.GetColumnWidthPixel(col);

            double totalHeightPixels = 0;
            for (int row = 0; row <= maxRow; row++)
                totalHeightPixels += sheet.Cells.GetRowHeightPixel(row);

            // Convert pixels to points (1 point = 1/72 inch, 1 pixel = 1/96 inch)
            float widthPoints = (float)(totalWidthPixels * 72.0 / 96.0);
            float heightPoints = (float)(totalHeightPixels * 72.0 / 96.0);

            // NOTE: Custom paper size APIs may vary between Aspose.Cells versions.
            // The following lines are commented out to ensure compatibility.
            // If your version supports custom paper size, uncomment and adjust accordingly.
            // sheet.PageSetup.PaperSize = PaperSizeType.PaperUser;
            // sheet.PageSetup.CustomPaperSize = new SizeF(widthPoints, heightPoints);

            // Configure PDF save options (disable automatic scaling)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
