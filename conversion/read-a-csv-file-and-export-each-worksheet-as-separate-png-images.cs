// Title: C# – Import CSV and Export Each Worksheet as PNG with Aspose.Cells
// Description: Load a CSV into an Aspose.Cells workbook, then iterate all worksheets and render each one to a separate PNG file using ImageOrPrintOptions (one page per sheet) and SheetRender.
// Keywords: Aspose.Cells CSV import C# | export worksheet to PNG | C# convert CSV to images | SheetRender PNG Aspose | one page per sheet image | batch worksheet image export
// Common Searches: Aspose.Cells read CSV and save sheets as PNG | C# render each Excel worksheet to separate PNG | How to export multiple worksheets to PNG using Aspose.Cells | Convert CSV data to PNG images with Aspose.Cells
// Developer Intent: Create PNG images for every worksheet after loading CSV data into an Aspose.Cells workbook.
// Use Cases: Generate visual snapshots of CSV‑derived reports for web or email embedding. | Automate archival of each worksheet as an image for documentation or comparison. | Batch‑process CSV files to produce per‑sheet PNGs for QA or review workflows.
// AI Prompts: Write C# code that reads a CSV with Aspose.Cells and saves each worksheet as an individual PNG, allowing a custom output folder and filename pattern. | Show how to modify ImageOrPrintOptions (DPI, margins, scaling) when exporting worksheets to PNG using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCsvToPng
{
    // Load a CSV into an Aspose.Cells workbook, then iterate all worksheets and render each one to a separate PNG file using ImageOrPrintOptions (one page per sheet) and SheetRender.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Import CSV data into the first worksheet (starting at cell A1)
            // Parameters: file name, delimiter, convert numeric data, first row, first column
            workbook.Worksheets[0].Cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Output directory for PNG images
            string outputDir = "output_images";
            Directory.CreateDirectory(outputDir);

            // Iterate through each worksheet and render it to a PNG image
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Configure image rendering options
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true // Render the whole sheet on a single page
                };

                // Create a SheetRender instance for the current worksheet
                SheetRender sheetRender = new SheetRender(sheet, options);

                // Build the output file name (e.g., Sheet_0.png, Sheet_1.png, ...)
                string imagePath = Path.Combine(outputDir, $"Sheet_{sheetIndex}.png");

                // Render the first (and only) page of the sheet to the PNG file
                sheetRender.ToImage(0, imagePath);

                // Release resources used by SheetRender
                sheetRender.Dispose();

                Console.WriteLine($"Worksheet {sheet.Name} exported to {imagePath}");
            }
        }
    }
}
