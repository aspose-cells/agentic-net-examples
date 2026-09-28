// Title: Export each worksheet page of an Excel file to separate 300 DPI TIFF images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, sets ImageOrPrintOptions to 300 dpi, and saves every worksheet page as an individual TIFF file. | Show how to use SheetRender together with ImageOrPrintOptions to generate high‑resolution TIFF images for all sheets and pages in a workbook. | Add error handling that checks for a missing input file and sanitizes worksheet names before creating TIFF output files.
// Common Searches: Aspose.Cells C# export Excel worksheets to 300 DPI TIFF files | How to render each sheet page as a separate TIFF using Aspose.Cells .NET | Set DPI for TIFF output when converting .xlsx with Aspose.Cells | Generate multi‑page TIFF from Excel workbook with high resolution using Aspose.Cells
// Tags: Aspose.Cells ImageOrPrintOptions DPI configuration | Aspose.Cells SheetRender multi‑page TIFF export | C# high‑resolution TIFF generation from Excel | safe worksheet name for file output Aspose.Cells | 300 DPI TIFF export from .xlsx using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing.Imaging;

// Loads an Excel workbook, configures ImageOrPrintOptions to 300 dpi, iterates through all worksheets and their pages with SheetRender, and saves each page as an individual TIFF file using sanitized sheet names, with basic error handling for missing files.
class WorkbookToTiff
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set up image rendering options for TIFF output at 300 DPI
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // ImageFormat is inferred from the output file extension, so we omit it
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Process each worksheet in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Render the worksheet using the defined options
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                // Save each page of the worksheet as a separate TIFF file
                for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                {
                    // Create a safe file name for the output
                    string safeSheetName = string.Concat(sheet.Name.Split(Path.GetInvalidFileNameChars()));
                    string outputFile = $"output_{safeSheetName}_page{pageIndex + 1}.tiff";

                    try
                    {
                        sheetRender.ToImage(pageIndex, outputFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to save page {pageIndex + 1} of sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine("Workbook has been rendered to TIFF images at 300 DPI.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
