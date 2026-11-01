// Title: Calculate pixel offset between two worksheet shapes and write the values to cells with Aspose.Cells for .NET
// Common Searches: how to get absolute pixel position of a shape in Aspose.Cells C# | compute horizontal and vertical offset between two Excel shapes using Aspose.Cells | store shape distance calculation result in a worksheet cell with Aspose.Cells | determine Euclidean distance between two shapes in an Excel file using C# | Aspose.Cells example for shape position to pixel conversion
// Tags: shape absolute pixel position Aspose.Cells | calculate shape offset Excel C# | write shape offset values to worksheet cells | euclidean distance between worksheet shapes | Aspose.Cells pixel conversion for shape coordinates | retrieve shape column width row height pixels | C# Aspose.Cells shape position example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, ensures at least two shapes exist, computes each shape's absolute X and Y positions in pixels by summing column widths and row heights, calculates horizontal, vertical, and Euclidean offsets, writes these values to cells A1, B1, and C1, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there are at least two shapes
            if (sheet.Shapes.Count < 2)
                throw new InvalidOperationException("The worksheet must contain at least two shapes.");

            Shape shape1 = sheet.Shapes[0];
            Shape shape2 = sheet.Shapes[1];

            // Helper to calculate absolute X position in pixels (ignores offset if not available)
            double GetAbsoluteX(Shape shape)
            {
                double x = 0;
                for (int col = 0; col < shape.UpperLeftColumn; col++)
                    x += sheet.Cells.GetColumnWidthPixel(col);
                // UpperLeftColumnOffset is not available in this version; assume zero offset
                return x;
            }

            // Helper to calculate absolute Y position in pixels (ignores offset if not available)
            double GetAbsoluteY(Shape shape)
            {
                double y = 0;
                for (int row = 0; row < shape.UpperLeftRow; row++)
                    y += sheet.Cells.GetRowHeightPixel(row);
                // UpperLeftRowOffset is not available in this version; assume zero offset
                return y;
            }

            // Absolute pixel positions for both shapes
            double x1 = GetAbsoluteX(shape1);
            double y1 = GetAbsoluteY(shape1);
            double x2 = GetAbsoluteX(shape2);
            double y2 = GetAbsoluteY(shape2);

            // Calculate offsets and Euclidean distance
            double offsetX = Math.Abs(x2 - x1);
            double offsetY = Math.Abs(y2 - y1);
            double distance = Math.Sqrt(offsetX * offsetX + offsetY * offsetY);

            // Store results in cells A1 (X offset), B1 (Y offset), C1 (distance)
            sheet.Cells["A1"].PutValue(offsetX);
            sheet.Cells["B1"].PutValue(offsetY);
            sheet.Cells["C1"].PutValue(distance);

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Results saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
