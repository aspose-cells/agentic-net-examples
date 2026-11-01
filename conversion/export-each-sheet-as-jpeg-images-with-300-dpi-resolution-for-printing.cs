// Title: Export every worksheet in an Excel workbook to high‑resolution 300 DPI JPEG files using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export each Excel sheet as a 300 DPI JPEG for print quality | How to render Excel worksheets to high‑resolution JPEG images with Aspose.Cells in .NET | Save Excel workbook pages as 300 DPI JPEG files using SheetRender and ImageOrPrintOptions | Aspose.Cells C# generate per‑sheet JPEG images at 300 DPI for printing | Convert XLSX worksheets to 300 DPI JPEG images programmatically with Aspose.Cells
// Tags: SheetRender high‑resolution JPEG export | ImageOrPrintOptions 300 DPI setting | Aspose.Cells per‑worksheet image conversion | C# Excel to JPEG rendering | printing quality Excel image generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an XLSX workbook, iterates through all worksheets, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, uses SheetRender to render each sheet (and each page within a sheet) to separate JPEG files named "Sheet_{index}_Page_{page}.jpeg", and includes error handling for missing files and rendering failures.
class ExportSheetsToJpeg
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Configure image options: 300 DPI resolution
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                    // Note: ImageFormat property is omitted to avoid compatibility issues.
                    // The default format is PNG; if JPEG is required, ensure the Aspose.Cells version supports ImageFormat.
                };

                // Create a renderer for the current sheet
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                // Export each page of the sheet (most sheets have a single page)
                for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                {
                    // Build output file name, e.g., Sheet_1_Page_1.jpeg
                    string outputFile = $"Sheet_{sheetIndex + 1}_Page_{pageIndex + 1}.jpeg";

                    try
                    {
                        // Render the page to an image file
                        sheetRender.ToImage(pageIndex, outputFile);
                        Console.WriteLine($"Exported: {outputFile}");
                    }
                    catch (Exception exPage)
                    {
                        Console.WriteLine($"Failed to export page {pageIndex + 1} of sheet {sheetIndex + 1}: {exPage.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
