using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class WorksheetToImageConverter
    {
        /// <summary>
        /// Loads an existing XLSX file, renders the specified worksheet to PNG images
        /// (one image per printed page) while preserving layout and visual fidelity.
        /// </summary>
        public static void Run()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Ensure the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // Choose the worksheet to render (e.g., the first worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Output image format
                ImageType = Aspose.Cells.Drawing.ImageType.Png,

                // Render each printed page as a separate image
                OnePagePerSheet = true,

                // Optional: increase resolution for higher quality
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Create a SheetRender instance for the selected worksheet
            SheetRender sheetRender = new SheetRender(worksheet, options);

            // Prepare output directory
            string outputDir = "output_images";
            Directory.CreateDirectory(outputDir);

            // Render each page of the worksheet to a separate PNG file
            for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
            {
                string imagePath = Path.Combine(outputDir, $"worksheet_page_{pageIndex}.png");

                // Render the page to the specified file
                sheetRender.ToImage(pageIndex, imagePath);

                Console.WriteLine($"Rendered page {pageIndex} to image: {imagePath}");
            }

            // Clean up resources
            sheetRender.Dispose();
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetToImageConverter.Run();
        }
    }
}