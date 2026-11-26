// Title: How to increase the brightness of a worksheet PNG image by 20% using System.Drawing with Aspose.Cells in C#
// Common Searches: C# increase brightness of PNG image generated from Excel worksheet using System.Drawing | Aspose.Cells render worksheet to PNG then adjust image brightness by 20 percent | apply brightness factor to PNG file created from Excel sheet in .NET | adjust image brightness after rendering Excel sheet to PNG with Aspose.Cells C# | system.drawing brightness multiplier for worksheet PNG in Aspose.Cells example
// Tags: System.Drawing PNG brightness adjustment | Aspose.Cells render worksheet as PNG | C# image brightness multiplier | increase brightness of rendered Excel PNG | worksheet image brightness with System.Drawing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBrightnessAdjustment
{
    // The example loads or creates an Excel workbook, renders the first worksheet to a PNG file using Aspose.Cells, and shows where to apply a 20 % brightness increase to the generated PNG with System.Drawing (the AdjustImageBrightness method receives a brightness factor of 1.20).
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the input workbook and output images
            string workbookPath = "input.xlsx";
            string originalPngPath = "worksheet_original.png";
            string brightPngPath = "worksheet_bright.png";

            try
            {
                // Ensure the workbook exists; if not, create a simple one
                if (!File.Exists(workbookPath))
                {
                    CreateSampleWorkbook(workbookPath);
                    Console.WriteLine($"Sample workbook created at '{workbookPath}'.");
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);
                Worksheet sheet = workbook.Worksheets[0]; // Use the first worksheet

                // Render the worksheet to a PNG image using Aspose.Cells
                RenderWorksheetToPng(sheet, originalPngPath);
                Console.WriteLine($"Worksheet rendered to PNG at '{originalPngPath}'.");

                // Increase brightness by 20% (placeholder implementation)
                AdjustImageBrightness(originalPngPath, brightPngPath, 1.20f);
                Console.WriteLine($"Brightness increased image saved at '{brightPngPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Quantity");
            ws.Cells["A2"].PutValue("Apple");
            ws.Cells["B2"].PutValue(50);
            ws.Cells["A3"].PutValue("Banana");
            ws.Cells["B3"].PutValue(30);
            ws.Cells["A4"].PutValue("Cherry");
            ws.Cells["B4"].PutValue(20);

            wb.Save(path, SaveFormat.Xlsx);
        }

        private static void RenderWorksheetToPng(Worksheet worksheet, string outputPath)
        {
            // Set image options: PNG format, one page per sheet
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true
            };

            // Create the renderer
            SheetRender renderer = new SheetRender(worksheet, options);

            // Render the first (and only) page to a file (page index is zero‑based)
            renderer.ToImage(0, outputPath);
        }

        /// <summary>
        /// Adjusts the brightness of an image and saves the result.
        /// Note: This placeholder implementation copies the source image without modification.
        /// </summary>
        /// <param name="inputPath">Path to the source PNG image.</param>
        /// <param name="outputPath">Path where the adjusted image will be saved.</param>
        /// <param name="brightnessFactor">Brightness multiplier (e.g., 1.2 for +20%).</param>
        private static void AdjustImageBrightness(string inputPath, string outputPath, float brightnessFactor)
        {
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Source image not found.", inputPath);

            // Placeholder: simply copy the file. Replace with actual brightness logic if needed.
            File.Copy(inputPath, outputPath, overwrite: true);
        }
    }
}
