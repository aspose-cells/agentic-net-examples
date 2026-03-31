using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsConversionDemo
{
    public class XlsxToTiffConverter
    {
        public static void Convert(string sourceXlsxPath, string destinationTiffPath)
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook(sourceXlsxPath);

            // Configure image rendering options for TIFF output
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                     // Set output format to TIFF
                TiffCompression = TiffCompression.CompressionLZW, // Use LZW compression for better quality
                HorizontalResolution = 300,                     // Set horizontal DPI
                VerticalResolution = 300,                       // Set vertical DPI
                OnePagePerSheet = true                          // Render each worksheet as a separate page in the multi‑page TIFF
            };

            // Create a workbook renderer with the configured options
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Render the entire workbook to a multi‑page TIFF file
            renderer.ToImage(destinationTiffPath);

            Console.WriteLine($"Workbook successfully converted to TIFF: {destinationTiffPath}");
        }

        // Example usage
        public static void Main()
        {
            string inputFile = "input.xlsx";      // Path to the source XLSX file
            string outputFile = "output.tiff";    // Desired path for the TIFF image

            Convert(inputFile, outputFile);
        }
    }
}