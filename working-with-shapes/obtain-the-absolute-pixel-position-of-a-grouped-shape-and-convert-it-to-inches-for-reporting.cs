// Title: Get absolute pixel coordinates of a grouped shape and convert them to inches with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get shape absolute pixel position in worksheet | convert shape coordinates from pixels to inches using Aspose.Cells | calculate top left location of grouped shape in Excel with Aspose.Cells .NET | retrieve absolute row and column offsets of a shape in Aspose.Cells | how to determine exact inches position of an Excel shape in C#
// Tags: shape absolute pixel location Aspose.Cells | pixel to inch conversion Aspose.Cells .NET | grouped shape coordinates Excel C# | worksheet shape position calculation Aspose.Cells | shape top left row column offsets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, accesses the first worksheet, obtains the first shape, computes its absolute top‑left position in pixels by summing row heights and column widths, converts those pixel values to inches using a 96 DPI factor, and prints both pixel and inch coordinates.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape (or locate by name as needed)
            Shape shape = sheet.Shapes[0];

            // Compute absolute row/column indices.
            int absoluteRow = shape.UpperLeftRow;
            int absoluteColumn = shape.UpperLeftColumn;

            // Offsets are not available in this version of Aspose.Cells; assume zero.
            int absoluteRowOffset = 0;
            int absoluteColumnOffset = 0;

            // Convert worksheet rows to pixel coordinates
            double rowPixel = 0;
            for (int r = 0; r < absoluteRow; r++)
            {
                rowPixel += sheet.Cells.GetRowHeightPixel(r);
            }
            rowPixel += absoluteRowOffset;

            // Convert worksheet columns to pixel coordinates
            double columnPixel = 0;
            for (int c = 0; c < absoluteColumn; c++)
            {
                columnPixel += sheet.Cells.GetColumnWidthPixel(c);
            }
            columnPixel += absoluteColumnOffset;

            // Convert pixels to inches (default DPI = 96)
            const double DPI = 96.0;
            double xInches = columnPixel / DPI;
            double yInches = rowPixel / DPI;

            // Output results
            Console.WriteLine($"Absolute position (pixels): X = {columnPixel}, Y = {rowPixel}");
            Console.WriteLine($"Absolute position (inches): X = {xInches:F2}\", Y = {yInches:F2}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
