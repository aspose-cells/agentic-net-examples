using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsWorksheetToPng
{
    public class Program
    {
        public static void Main()
        {
            // Load the workbook (replace with your actual file path)
            string workbookPath = "input.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // Ensure the output directory exists
            string outputDir = "output";
            Directory.CreateDirectory(outputDir);

            // Iterate through each worksheet and save it as a separate PNG file
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet worksheet = workbook.Worksheets[sheetIndex];

                // Configure image rendering options
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true // render the whole sheet on a single page
                };

                // Create a SheetRender for the current worksheet
                SheetRender sheetRender = new SheetRender(worksheet, options);

                // Build the output file name (e.g., Sheet_1.png, Sheet_2.png, ...)
                string outputPath = Path.Combine(outputDir, $"Sheet_{sheetIndex + 1}.png");

                // Render the first (and only) page of the sheet to a PNG file
                sheetRender.ToImage(0, outputPath);

                // Release resources used by SheetRender
                sheetRender.Dispose();

                Console.WriteLine($"Worksheet '{worksheet.Name}' saved as PNG to: {outputPath}");
            }
        }
    }
}