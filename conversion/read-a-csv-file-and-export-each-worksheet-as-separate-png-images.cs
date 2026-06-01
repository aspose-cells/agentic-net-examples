using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvToPng
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Temporary XLSX file that will hold the converted CSV data
            string tempXlsxPath = "temp_converted.xlsx";

            // Convert CSV to XLSX using the provided ConversionUtility rule
            ConversionUtility.Convert(csvPath, tempXlsxPath);

            // Load the converted workbook (creation rule)
            Workbook workbook = new Workbook(tempXlsxPath);

            // Configure image rendering options (PNG, one page per sheet)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true
            };

            // Ensure output directory exists
            string outputDir = "output_images";
            Directory.CreateDirectory(outputDir);

            // Iterate through each worksheet and render it to a separate PNG file
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Create SheetRender for the current worksheet (constructor rule)
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                // Build output file name (e.g., Sheet_1.png, Sheet_2.png, ...)
                string outputPath = Path.Combine(outputDir, $"Sheet_{sheetIndex + 1}.png");

                // Render the first (and only) page of the sheet to a PNG file (ToImage overload rule)
                sheetRender.ToImage(0, outputPath);

                // Release resources used by SheetRender
                sheetRender.Dispose();
            }

            // Clean up the temporary XLSX file
            if (File.Exists(tempXlsxPath))
            {
                File.Delete(tempXlsxPath);
            }

            Console.WriteLine("All worksheets have been exported as PNG images.");
        }
    }
}