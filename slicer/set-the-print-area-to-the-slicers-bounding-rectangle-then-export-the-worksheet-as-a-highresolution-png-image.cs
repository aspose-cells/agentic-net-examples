// Title: C# – Set Print Area to a Slicer’s Bounding Rectangle and Export as High‑Resolution PNG with Aspose.Cells
// Description: Loads an Excel file, reads the first slicer’s shape coordinates, sets the worksheet print area to that bounding range, and renders the area to a 300 DPI PNG using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# slicer print area | export slicer to PNG | high resolution worksheet image | set print area from slicer shape | Aspose.Cells ImageOrPrintOptions | C# Excel slicer export | 300 DPI PNG Aspose.Cells
// Common Searches: Aspose.Cells set print area from slicer | C# export slicer region as PNG | high DPI worksheet image Aspose.Cells | get slicer shape bounds Aspose.Cells | render slicer area to image .NET
// Developer Intent: Define the worksheet print area based on a slicer’s bounding rectangle and generate a high‑resolution PNG of that area.
// Use Cases: Create a snapshot of a slicer for dashboards or reports. | Produce printable PNGs of filtered data views for documentation. | Automate generation of high‑quality images for web or PDF embedding.
// AI Prompts: Generate C# code that sets the print area to a slicer’s shape bounds and saves a 300 DPI PNG with Aspose.Cells. | Explain how to retrieve a slicer’s UpperLeftRow/Column and LowerRightRow/Column to define a print area in Aspose.Cells. | Show how to loop through all slicers in a workbook and export each to a separate high‑resolution PNG file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerExport
{
    // Loads an Excel file, reads the first slicer’s shape coordinates, sets the worksheet print area to that bounding range, and renders the area to a 300 DPI PNG using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "slicer_area.png";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                    return;
                }

                // Load the workbook that contains a slicer
                Workbook workbook = new Workbook(inputFile);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one slicer
                if (worksheet.Slicers.Count == 0)
                {
                    Console.WriteLine("Error: No slicers found in the worksheet.");
                    return;
                }

                // Get the first slicer
                Slicer slicer = worksheet.Slicers[0];

                // Retrieve the slicer's bounding rectangle via its shape
                Shape slicerShape = worksheet.Shapes[slicer.Name];
                int startRow = slicerShape.UpperLeftRow;
                int endRow = slicerShape.LowerRightRow;
                int startColumn = slicerShape.UpperLeftColumn;
                int endColumn = slicerShape.LowerRightColumn;

                // Convert cell indices to A1 style addresses
                string startCell = CellsHelper.CellIndexToName(startRow, startColumn);
                string endCell = CellsHelper.CellIndexToName(endRow, endColumn);

                // Set the worksheet's print area to the slicer's bounding rectangle
                worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

                // Configure high‑resolution image options
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true,
                    HorizontalResolution = 300, // DPI
                    VerticalResolution = 300    // DPI
                };

                // Render the worksheet (print area only) to a PNG image
                SheetRender sheetRender = new SheetRender(worksheet, options);
                sheetRender.ToImage(0, outputFile);

                Console.WriteLine($"Export completed: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
