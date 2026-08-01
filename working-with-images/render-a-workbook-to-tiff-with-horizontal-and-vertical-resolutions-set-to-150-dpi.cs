// Title: Export an Aspose.Cells Workbook to a Multi‑Page TIFF at 150 DPI (C#)
// Description: Demonstrates how to create or load a workbook, set ImageOrPrintOptions to TIFF format, configure HorizontalResolution and VerticalResolution to 150 DPI, and use WorkbookRender to generate a multi‑page TIFF file in C#.
// Keywords: Aspose.Cells | C# | .NET | TIFF export | 150 DPI | ImageOrPrintOptions | WorkbookRender | multi‑page TIFF | Excel to TIFF | set DPI Aspose.Cells
// Common Searches: Aspose.Cells export workbook to TIFF C# | set horizontal and vertical DPI for TIFF in Aspose.Cells | render Excel as multi‑page TIFF with 150 DPI | ImageOrPrintOptions DPI settings Aspose.Cells | C# create high‑resolution TIFF from spreadsheet
// Developer Intent: Generate a TIFF image of the entire workbook with both horizontal and vertical resolution fixed at 150 DPI.
// Use Cases: Archiving spreadsheets as high‑resolution printable TIFFs for compliance records. | Embedding multi‑page TIFFs in reports or document management systems that require a specific DPI. | Converting Excel data to a format compatible with legacy imaging workflows that accept only TIFF files.
// AI Prompts: Modify the sample to accept a DPI value from the user and apply it to both resolutions. | Show how to render a single worksheet to a 150 DPI TIFF instead of the whole workbook. | Add TIFF compression options while preserving the 150 DPI setting.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffDemo
{
    // Demonstrates how to create or load a workbook, set ImageOrPrintOptions to TIFF format, configure HorizontalResolution and VerticalResolution to 150 DPI, and use WorkbookRender to generate a multi‑page TIFF file in C#.
    public class RenderWorkbookToTiff
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data to demonstrate the rendering
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF Rendering Demo");
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Specify that the output format is TIFF
                ImageType = ImageType.Tiff,
                // Set the desired DPI for both horizontal and vertical dimensions
                HorizontalResolution = 150,
                VerticalResolution = 150
            };

            // Create a workbook renderer with the configured options
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Render the entire workbook to a multi‑page TIFF file
            string outputPath = "WorkbookOutput_150dpi.tiff";
            renderer.ToImage(outputPath);

            Console.WriteLine($"Workbook successfully rendered to TIFF at {outputPath} with 150 DPI resolution.");
        }
    }
}
