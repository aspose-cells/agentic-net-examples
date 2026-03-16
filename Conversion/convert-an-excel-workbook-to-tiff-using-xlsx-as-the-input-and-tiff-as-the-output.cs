using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsConversionDemo
{
    public class ExcelToTiffConverter
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired path for the output TIFF file
            string outputPath = "output.tiff";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet (you can choose any worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure rendering options (optional settings)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // Render each worksheet on a single page
                OnePagePerSheet = true
            };

            // Create a SheetRender object using the worksheet and options
            SheetRender sheetRenderer = new SheetRender(worksheet, renderOptions);

            // Render the entire worksheet to a multi‑page TIFF file
            sheetRenderer.ToTiff(outputPath);

            Console.WriteLine($"Workbook '{sourcePath}' successfully converted to TIFF at '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelToTiffConverter.Run();
        }
    }
}