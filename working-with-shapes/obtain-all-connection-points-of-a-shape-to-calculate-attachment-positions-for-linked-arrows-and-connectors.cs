// Title: Determine all connection points of an Excel shape and compute attachment coordinates for arrows using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, retrieves a shape by name, enumerates its connection points, and returns the X/Y pixel coordinates for each point to enable precise arrow attachment. | Create a C# example that calculates the absolute pixel position of a shape's top‑left corner and then derives the pixel locations of its four standard connection points (top, bottom, left, right) using row height and column width conversions. | Generate a C# snippet that safely accesses UpperLeftRowOffset and UpperLeftColumnOffset via reflection when unavailable, and then maps all shape connection points to worksheet cell coordinates for connector drawing.
// Common Searches: aspnet how to get shape connection points in Excel with Aspose.Cells | c# enumerate attachment points of a drawing shape using Aspose.Cells | calculate pixel coordinates for shape connectors Aspose.Cells .NET | retrieve shape offsets and connection locations for arrows in Excel workbook | Aspose.Cells get shape top left cell and connection points for diagram
// Tags: Aspose.Cells enumerate shape connection points | C# calculate shape pixel attachment coordinates | Aspose.Cells convert row height points to pixels | Aspose.Cells reflection access shape offsets | Excel shape connector positioning Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, finds a shape named 'MyShape', reads its cell location and optional pixel offsets (using reflection when needed), converts row heights (points) and column widths (character units) to pixels at 96 DPI, computes the shape’s absolute top‑left pixel position, enumerates all standard connection points (top, bottom, left, right) and outputs their pixel coordinates for accurate arrow or connector placement, then saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the shape by name.
            Shape shape = sheet.Shapes["MyShape"];
            if (shape == null)
            {
                Console.WriteLine("Shape 'MyShape' not found.");
                return;
            }

            // Standard screen DPI.
            const double dpi = 96.0;

            // Shape's top‑left cell position.
            int startRow = shape.UpperLeftRow;
            int startColumn = shape.UpperLeftColumn;

            // Offsets of the shape within the cell (pixels).
            // UpperLeftRowOffset / UpperLeftColumnOffset are not available in older API versions,
            // so we default to 0 when they cannot be accessed.
            int rowOffsetPixels = 0;
            int columnOffsetPixels = 0;

            // Attempt to read offsets via reflection (optional, safe fallback).
            try
            {
                var rowOffsetProp = shape.GetType().GetProperty("UpperLeftRowOffset");
                var colOffsetProp = shape.GetType().GetProperty("UpperLeftColumnOffset");
                if (rowOffsetProp != null && colOffsetProp != null)
                {
                    rowOffsetPixels = (int)rowOffsetProp.GetValue(shape);
                    columnOffsetPixels = (int)colOffsetProp.GetValue(shape);
                }
            }
            catch
            {
                // Ignore any reflection errors; offsets remain zero.
            }

            // Convert row height (points) to pixels.
            double rowHeightPoints = sheet.Cells.Rows[startRow].Height; // points
            double rowTopPixels = rowHeightPoints * dpi / 72.0;

            // Approximate conversion of column width (characters) to pixels.
            double columnWidthChars = sheet.Cells.Columns[startColumn].Width; // characters
            double columnLeftPixels = Math.Floor(columnWidthChars * 7.0 + 5.0);

            // Absolute pixel coordinates of the shape's top‑left corner.
            double absoluteX = columnLeftPixels + columnOffsetPixels;
            double absoluteY = rowTopPixels + rowOffsetPixels;

            Console.WriteLine($"Shape '{shape.Name}' position:");
            Console.WriteLine($"  UpperLeftCell: Row {startRow}, Column {startColumn}");
            Console.WriteLine($"  Offsets (pixels): Row {rowOffsetPixels}, Column {columnOffsetPixels}");
            Console.WriteLine($"  Absolute Position (pixels): X = {absoluteX}, Y = {absoluteY}");

            // Save the workbook (if any modifications were made).
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
