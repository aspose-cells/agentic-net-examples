// Title: C# – Add a Rectangle Shape, Retrieve Connection Points, and Map to Excel Cell Addresses with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, inserts a rectangle shape, extracts its connection points via GetConnectionPoints(), converts each X/Y coordinate to the nearest column and row using default column width and row height, clamps the indices to worksheet limits, and prints the corresponding cell addresses before saving the file.
// Keywords: Aspose.Cells C# shape connection points | GetConnectionPoints Aspose.Cells | map shape points to Excel cells | convert shape coordinates to cell address | Aspose.Cells shape anchor mapping | C# Excel shape example
// Common Searches: Aspose.Cells get shape connection points C# | map shape connection points to worksheet cells | convert shape coordinates to Excel cell address .NET | retrieve and display shape connection points Aspose.Cells | C# example for shape anchor mapping in Excel
// Developer Intent: Add a shape to a worksheet, obtain its connection points, and determine the exact cell addresses that correspond to those points.
// Use Cases: Generate a documentation sheet that lists each shape's connection points alongside their cell references. | Create a mapping table for downstream processes that need to align shape anchors with spreadsheet grid locations. | Validate shape placement by comparing connection point addresses with expected cell ranges.
// AI Prompts: Write C# code using Aspose.Cells to add a shape, retrieve its connection points, and precisely convert each point to the exact cell address. | Explain how to calculate row and column indices from shape point coordinates using default column width and row height in Aspose.Cells. | Suggest a more accurate method for mapping shape connection points to cells, leveraging worksheet column width and row height properties.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This Aspose.Cells for .NET example creates a workbook, inserts a rectangle shape, extracts its connection points via GetConnectionPoints(), converts each X/Y coordinate to the nearest column and row using default column width and row height, clamps the indices to worksheet limits, and prints the corresponding cell addresses before saving the file.
class ShapeConnectionPointsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, upper left row offset, upper left column offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,                        // upper left row
                1,                        // upper left column
                0,                        // upper left row offset (in pixels)
                0,                        // upper left column offset (in pixels)
                80,                       // height (in points)
                150);                     // width (in points)

            // Retrieve the connection points of the shape
            float[][] points = shape.GetConnectionPoints();

            Console.WriteLine("Connection Points and Corresponding Cell Addresses:");
            for (int i = 0; i < points.Length; i++)
            {
                float x = points[i][0];
                float y = points[i][1];

                // Approximate conversion from point coordinates to column/row indices.
                // 1 point = 1/72 inch. Assuming default column width (~64 pixels) and row height (~15 points).
                int columnIndex = (int)Math.Floor(x / 72.0 * 8.43);
                int rowIndex = (int)Math.Floor(y / 15.0);

                // Clamp indices to worksheet bounds
                columnIndex = Math.Max(0, Math.Min(sheet.Cells.MaxColumn, columnIndex));
                rowIndex = Math.Max(0, Math.Min(sheet.Cells.MaxRow, rowIndex));

                // Get the cell address (e.g., "B3")
                string cellAddress = sheet.Cells[rowIndex, columnIndex].Name;

                Console.WriteLine($"Point {i + 1}: X={x}, Y={y} => Cell {cellAddress}");
            }

            // Save the workbook (optional, for verification)
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeConnectionPointsDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
