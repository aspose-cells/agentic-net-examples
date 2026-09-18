// Title: Export a worksheet by zero‑based index to a 200 DPI JPEG image using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, selects a worksheet at index 2, and saves it as a JPEG image at 200 DPI. | Demonstrate how to configure ImageOrPrintOptions for 200 DPI horizontal and vertical resolution and use SheetRender to export a single worksheet page to JPEG. | Create a C# snippet that checks the worksheet index range, creates the output folder if missing, and handles errors while rendering the sheet to a high‑resolution JPEG.
// Common Searches: Aspose.Cells C# export specific worksheet to JPEG with 200 DPI resolution | How to render an Excel sheet by index as a high‑resolution image in .NET | Set horizontal and vertical DPI when converting Excel worksheet to JPEG using Aspose.Cells | C# code to save the third sheet of an Excel workbook as a 200 DPI JPEG file | Validate worksheet index before exporting to image with Aspose.Cells
// Tags: Aspose.Cells SheetRender JPEG export | ImageOrPrintOptions DPI configuration | C# high‑resolution Excel to image conversion | Check worksheet range before rendering | Create output directory for image export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads a workbook, verifies the source file and output folder, ensures the requested zero‑based worksheet index is valid, sets ImageOrPrintOptions to 200 DPI horizontal and vertical resolution, creates a SheetRender for the selected worksheet, and renders the first page to a JPEG file.
class WorksheetToJpeg
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string sourceFile = @"C:\Input\Workbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            // Index of the worksheet to render (0‑based)
            int worksheetIndex = 2; // example: third worksheet

            // Path for the output JPEG image
            string outputFile = @"C:\Output\Worksheet3.jpg";

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputFile);
            if (string.IsNullOrEmpty(outputDir))
            {
                Console.WriteLine("Invalid output path.");
                return;
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourceFile);

            // Ensure the requested index is within range
            if (worksheetIndex < 0 || worksheetIndex >= workbook.Worksheets.Count)
            {
                Console.WriteLine("Worksheet index out of range.");
                return;
            }

            // Configure rendering options with 200 DPI resolution
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                HorizontalResolution = 200,
                VerticalResolution = 200
                // Image format is inferred from the output file extension (JPEG)
                // OnePagePerSheet = true // uncomment to fit the whole sheet on one page
            };

            // Create a SheetRender object for the target worksheet
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[worksheetIndex], options);

            // Render the first (and only) page of the worksheet to an image
            sheetRender.ToImage(0, outputFile);

            Console.WriteLine($"Worksheet {worksheetIndex} rendered to JPEG at {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
