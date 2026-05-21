using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffSplitDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Configure image options for TIFF rendering
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,               // Output format
                TiffCompression = TiffCompression.CompressionLZW,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Create a renderer for the entire workbook
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Determine how many pages the workbook will generate
            int totalPages = renderer.PageCount;

            // Ensure the output directory exists
            string outputDir = "TiffPages";
            Directory.CreateDirectory(outputDir);

            // Render each page to a separate single‑page TIFF file
            for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
            {
                string pageFile = Path.Combine(outputDir, $"Page_{pageIndex + 1}.tiff");
                renderer.ToImage(pageIndex, pageFile);
                Console.WriteLine($"Saved page {pageIndex + 1} to {pageFile}");
            }

            // Clean up resources
            renderer.Dispose();
        }
    }
}