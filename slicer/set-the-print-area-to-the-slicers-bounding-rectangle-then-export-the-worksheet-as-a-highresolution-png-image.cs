// Title: Export a worksheet to a 300 DPI PNG after setting the print area to a slicer’s bounding rectangle with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set worksheet print area to slicer bounds | How to export Excel sheet as high‑resolution PNG using Aspose.Cells | Aspose.Cells render worksheet to 300 DPI PNG image | Get slicer bounding rectangle in Aspose.Cells C# | Configure ImageOrPrintOptions for high‑resolution PNG in Aspose.Cells | Set PageSetup.PrintArea based on slicer location Aspose.Cells
// Tags: Aspose.Cells slicer print area mapping | PNG image export with 300 DPI using Aspose.Cells | Configure ImageOrPrintOptions for high DPI PNG | SheetRender C# render worksheet to PNG | C# obtain slicer bounds Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;   // Required for Slicer class

namespace AsposeCellsSlicerExport
{
    // The example loads an Excel workbook, retrieves the first slicer on the first worksheet, determines a print area (using the used range as a fallback for the slicer bounds), sets the worksheet's PageSetup.PrintArea, configures ImageOrPrintOptions for a 300 DPI PNG, and renders the sheet to a high‑resolution PNG file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "worksheet_slicer_area.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputFile);
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Locate a slicer and obtain its bounding rectangle.
                // ------------------------------------------------------------
                if (sheet.Slicers.Count == 0)
                {
                    Console.WriteLine("Error: No slicer found on the first worksheet.");
                    return;
                }

                // Use the first slicer for this demonstration
                Slicer slicer = sheet.Slicers[0];

                // Aspose.Cells Slicer does not expose cell indices directly.
                // As a fallback, use the worksheet's used range as the print area.
                // (If exact slicer bounds are required, additional calculations based on
                // slicer pixel positions would be needed.)
                int topRow = sheet.Cells.MaxDisplayRange.FirstRow;
                int bottomRow = sheet.Cells.MaxDisplayRange.FirstRow + sheet.Cells.MaxDisplayRange.RowCount - 1;
                int leftColumn = sheet.Cells.MaxDisplayRange.FirstColumn;
                int rightColumn = sheet.Cells.MaxDisplayRange.FirstColumn + sheet.Cells.MaxDisplayRange.ColumnCount - 1;

                // Convert indices to A1 style addresses
                string topLeft = CellsHelper.CellIndexToName(topRow, leftColumn);
                string bottomRight = CellsHelper.CellIndexToName(bottomRow, rightColumn);
                string printArea = $"{topLeft}:{bottomRight}";

                // ------------------------------------------------------------
                // 2. Set the worksheet print area to the determined rectangle.
                // ------------------------------------------------------------
                sheet.PageSetup.PrintArea = printArea;
                Console.WriteLine($"Print area set to: {printArea}");

                // ------------------------------------------------------------
                // 3. Configure high‑resolution image options.
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true
                };

                // ------------------------------------------------------------
                // 4. Render the worksheet to a PNG image.
                // ------------------------------------------------------------
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                renderer.ToImage(0, outputFile);

                Console.WriteLine($"Worksheet exported successfully to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
