// Title: Create a multi‑page TIFF from an Excel workbook with 150 DPI resolution using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel to TIFF with 150 DPI horizontal and vertical resolution | How to set custom DPI when rendering a workbook to multi‑page TIFF in Aspose.Cells | Aspose.Cells ImageOrPrintOptions set HorizontalResolution and VerticalResolution for TIFF output | Render all worksheets to a single TIFF file using Aspose.Cells .NET | Apply LZW compression while exporting Excel to TIFF with Aspose.Cells
// Tags: Aspose.Cells render workbook to TIFF | ImageOrPrintOptions DPI settings | WorkbookRender multi-page TIFF export | C# set TiffCompression LZW Aspose.Cells | Excel to TIFF conversion with custom resolution

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffRender
{
    // The example creates or loads an Excel workbook, populates it with data and a formula, then configures ImageOrPrintOptions to use TIFF format, sets HorizontalResolution and VerticalResolution to 150 DPI, applies LZW compression, and uses WorkbookRender to export all worksheets into a single multi‑page TIFF file, handling output folder creation and proper disposal of resources.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Load or create a workbook.
                // -----------------------------------------------------------------
                // If you have an existing Excel file, replace the below line with:
                // Workbook workbook = new Workbook("input.xlsx");
                Workbook workbook = new Workbook();

                // Add some sample data to demonstrate the rendering.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["C2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["C3"].PutValue(0.3);

                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B4"].PutValue(60);
                sheet.Cells["C4"].PutValue(0.4);

                // Add a simple total formula.
                sheet.Cells["B5"].PutValue("Total:");
                sheet.Cells["C5"].Formula = "=SUM(C2:C4)";

                // Calculate formulas so that the result is stored.
                workbook.CalculateFormula();

                // -----------------------------------------------------------------
                // 2. Configure image rendering options for TIFF with 150 DPI.
                // -----------------------------------------------------------------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Tiff, // Output format
                    HorizontalResolution = 150,                     // 150 DPI horizontally
                    VerticalResolution = 150,                       // 150 DPI vertically
                    TiffCompression = TiffCompression.CompressionLZW // Optional: LZW compression
                };

                // -----------------------------------------------------------------
                // 3. Render the entire workbook to a single multi‑page TIFF file.
                // -----------------------------------------------------------------
                string outputPath = "WorkbookOutput.tiff";

                // Ensure the directory for the output file exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                WorkbookRender renderer = null;
                try
                {
                    renderer = new WorkbookRender(workbook, renderOptions);
                    renderer.ToImage(outputPath); // Renders all worksheets into one TIFF
                    Console.WriteLine($"Workbook successfully rendered to TIFF at: {outputPath}");
                }
                finally
                {
                    // Dispose the renderer if it was created.
                    renderer?.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during rendering:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
