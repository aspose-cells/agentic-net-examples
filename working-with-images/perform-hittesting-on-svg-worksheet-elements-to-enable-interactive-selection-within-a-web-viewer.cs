// Title: How to perform hit‑testing on an SVG worksheet exported with Aspose.Cells for .NET to identify clicked Excel cells
// Common Searches: C# determine Excel cell from mouse click coordinates on SVG rendered by Aspose.Cells | Aspose.Cells hit test SVG sheet to get cell address based on pixel position | map pixel X Y to worksheet cell using GetRowHeightPixel and GetColumnWidthPixel in .NET | interactive Excel cell selection in a web viewer using Aspose.Cells SVG export | calculate cell name from click location on exported SVG sheet in C#
// Tags: SVG worksheet hit testing Aspose.Cells | convert pixel position to cell name .NET | GetRowHeightPixel GetColumnWidthPixel methods | SheetRender export worksheet to SVG C# | web viewer cell selection with Aspose.Cells SVG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, renders the first worksheet to an SVG file with SheetRender, and then determines which cell corresponds to given pixel coordinates by iterating over row heights and column widths using GetRowHeightPixel and GetColumnWidthPixel, returning the cell address or "None" when the click falls outside the used range.
class SvgHitTesting
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Export the worksheet to SVG using SheetRender
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg,
                OnePagePerSheet = true
            };
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToImage(0, "Sheet1.svg");

            // Example hit‑test: suppose the user clicked at pixel (150, 80) on the SVG.
            int clickX = 150; // X coordinate in pixels relative to the SVG image
            int clickY = 80;  // Y coordinate in pixels relative to the SVG image

            string cellAddress = HitTestCell(sheet, clickX, clickY);
            Console.WriteLine($"Clicked cell: {cellAddress}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Returns the address of the cell located at the given pixel coordinates.
    // If the coordinates are outside any cell, returns "None".
    public static string HitTestCell(Worksheet sheet, int pixelX, int pixelY)
    {
        try
        {
            // Determine the used range limits
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            double accumulatedY = 0;
            for (int row = 0; row <= maxRow; row++)
            {
                double rowHeight = sheet.Cells.GetRowHeightPixel(row);
                if (pixelY < accumulatedY + rowHeight)
                {
                    double accumulatedX = 0;
                    for (int col = 0; col <= maxCol; col++)
                    {
                        double colWidth = sheet.Cells.GetColumnWidthPixel(col);
                        if (pixelX < accumulatedX + colWidth)
                        {
                            Cell cell = sheet.Cells[row, col];
                            return cell.Name;
                        }
                        accumulatedX += colWidth;
                    }
                    break; // X is beyond the last column
                }
                accumulatedY += rowHeight;
            }

            return "None";
        }
        catch
        {
            // In case of unexpected errors, return "None"
            return "None";
        }
    }
}
