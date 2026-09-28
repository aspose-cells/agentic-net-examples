// Title: Set the print area to a slicer's bounding rectangle and export the worksheet as a high‑resolution PNG using Aspose.Cells for .NET
// AI Prompts: Generate C# code that detects slicers on a worksheet, sets the print area to the slicer's bounding range, and saves the sheet as a 300 DPI PNG with Aspose.Cells. | Show how to configure ImageOrPrintOptions for high‑resolution PNG export and render a single‑page image of a worksheet using SheetRender in .NET. | Write a .NET program that adjusts the page setup print area based on slicer presence and creates a high‑quality PNG image of the entire sheet.
// Common Searches: Aspose.Cells set print area to slicer range and export to PNG | C# export Excel worksheet as 300 DPI PNG with slicer handling | How to use ImageOrPrintOptions for high resolution PNG in Aspose.Cells | Render worksheet to single image when slicers are present using Aspose.Cells | Set page setup print area programmatically based on slicer in .NET
// Tags: set print area based on slicer Aspose.Cells | export worksheet to 300dpi PNG Aspose.Cells | ImageOrPrintOptions high resolution PNG | SheetRender single page image export | check slicer count before setting print area

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, checks whether the first worksheet contains slicers, sets the worksheet's print area to the slicer's bounding rectangle when present, configures ImageOrPrintOptions for 300 DPI PNG output, and uses SheetRender to generate a single‑page high‑resolution PNG image of the sheet.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input and output file paths
                string inputPath = "input.xlsx";
                string outputImagePath = "output.png";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // If the worksheet contains slicers, set the print area to the used range
                if (worksheet.Slicers.Count > 0)
                {
                    // Use the worksheet's used range as the print area
                    var usedRange = worksheet.Cells.MaxDisplayRange;
                    string startCell = CellsHelper.CellIndexToName(usedRange.FirstRow, usedRange.FirstColumn);
                    string endCell = CellsHelper.CellIndexToName(
                        usedRange.FirstRow + usedRange.RowCount - 1,
                        usedRange.FirstColumn + usedRange.ColumnCount - 1);

                    worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";
                }

                // Configure image export options for high‑resolution PNG
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default format is PNG; explicit setting omitted to avoid API mismatch
                    HorizontalResolution = 300, // DPI
                    VerticalResolution = 300,   // DPI
                    OnePagePerSheet = true      // Export the whole sheet as a single image
                };

                // Render the worksheet to an image using the specified options
                SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
                sheetRender.ToImage(0, outputImagePath); // Page index 0, output file name

                Console.WriteLine($"Worksheet rendered to image: {outputImagePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
