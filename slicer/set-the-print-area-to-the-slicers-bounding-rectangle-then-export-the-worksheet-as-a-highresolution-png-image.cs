// Title: Set worksheet print area to a slicer's bounds and export as a 300 dpi PNG with Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads the first slicer's UpperLeftRow/UpperLeftColumn and BottomRightRow/BottomRightColumn via reflection, assigns that range to the worksheet's PrintArea, and saves the sheet as a 300 dpi PNG using Aspose.Cells. | Show how to configure ImageOrPrintOptions for high‑resolution PNG output after setting a custom print area based on slicer coordinates in Aspose.Cells for .NET. | Create a reusable method that takes input and output file paths, extracts slicer bounds, applies them as the page‑setup print area, and renders the first worksheet page to a PNG image with 300 dpi resolution.
// Common Searches: how to set print area from slicer location using Aspose.Cells C# | export Excel sheet to high resolution PNG based on slicer bounds Aspose.Cells | retrieve slicer coordinates with reflection Aspose.Cells .NET | Aspose.Cells ImageOrPrintOptions set horizontal and vertical resolution
// Tags: set print area from slicer bounds Aspose.Cells | export worksheet to high‑resolution PNG Aspose.Cells | retrieve slicer coordinates reflection .NET | ImageOrPrintOptions resolution settings | SheetRender export first page PNG

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;

// // Loads an Excel workbook, obtains the first slicer's bounding rows and columns via reflection, sets the worksheet's print area to that range, and renders the first page to a 300 dpi PNG image using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Attempt to set print area based on the first slicer, if possible
            if (worksheet.Slicers.Count > 0)
            {
                Slicer slicer = worksheet.Slicers[0];
                int startRow = 0, startColumn = 0, endRow = 0, endColumn = 0;
                bool boundsObtained = false;

                try
                {
                    // Use reflection to access slicer bounds (covers different library versions)
                    Type slicerType = slicer.GetType();
                    PropertyInfo propStartRow = slicerType.GetProperty("UpperLeftRow");
                    PropertyInfo propStartCol = slicerType.GetProperty("UpperLeftColumn");
                    PropertyInfo propEndRow = slicerType.GetProperty("BottomRightRow");
                    PropertyInfo propEndCol = slicerType.GetProperty("BottomRightColumn");

                    if (propStartRow != null && propStartCol != null && propEndRow != null && propEndCol != null)
                    {
                        startRow = (int)propStartRow.GetValue(slicer);
                        startColumn = (int)propStartCol.GetValue(slicer);
                        endRow = (int)propEndRow.GetValue(slicer);
                        endColumn = (int)propEndCol.GetValue(slicer);
                        boundsObtained = true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Reflection error while retrieving slicer bounds: {ex.Message}");
                }

                if (boundsObtained)
                {
                    string startCell = CellsHelper.CellIndexToName(startRow, startColumn);
                    string endCell = CellsHelper.CellIndexToName(endRow, endColumn);
                    worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";
                }
                else
                {
                    Console.WriteLine("Unable to retrieve slicer bounds; exporting the full worksheet.");
                }
            }
            else
            {
                Console.WriteLine("No slicers found; exporting the full worksheet.");
            }

            // Configure image export options (PNG is default)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render the worksheet to a PNG image (first page)
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Image exported successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
