// Title: C# – Add a Rectangle Shape, Retrieve Its Connection Points, and Map Them to Excel Cell Addresses with Aspose.Cells
// Description: This example creates a new workbook, inserts a rectangle shape on the first worksheet, calls GetConnectionPoints() to obtain the shape's anchor points, converts each X/Y coordinate from points to column and row indices using column‑width and row‑height pixel helpers, translates the indices to standard cell names via CellsHelper, prints the mapping, and saves the file.
// Keywords: Aspose.Cells C# shape connection points | map shape coordinates to Excel cells | GetConnectionPoints Aspose.Cells | convert shape points to cell address | rectangle shape Aspose.Cells example
// Common Searches: Aspose.Cells get shape connection points C# | convert shape point to Excel cell name | map rectangle anchor to cell address Aspose | C# Aspose.Cells shape geometry to cell reference | retrieve and document shape connection points
// Developer Intent: Add a shape, extract its connection points, and translate those points into the corresponding Excel cell references.
// Use Cases: Document each anchor point of a diagram shape alongside its cell location for design reviews. | Programmatically align data cells with shape anchors when generating dynamic reports. | Store shape geometry in a worksheet to enable later reconstruction or positional analysis.
// AI Prompts: Generate C# code using Aspose.Cells that adds a rectangle shape, reads its connection points, and outputs the nearest cell address for each point. | Explain how to convert shape point coordinates to Excel cell names using column‑width and row‑height pixel helpers in Aspose.Cells. | Suggest a technique for more precise mapping of connection points when columns and rows have variable sizes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.IO;

// This example creates a new workbook, inserts a rectangle shape on the first worksheet, calls GetConnectionPoints() to obtain the shape's anchor points, converts each X/Y coordinate from points to column and row indices using column‑width and row‑height pixel helpers, translates the indices to standard cell names via CellsHelper, prints the mapping, and saves the file.
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
            // Parameters: upper left row, upper left column, upper left offsetX, offsetY, width, height
            Shape shape = sheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 80);

            // Retrieve the connection points of the shape
            float[][] points = shape.GetConnectionPoints();

            // Document each connection point and map it to an approximate cell address
            Console.WriteLine("Connection Points and Approximate Cell Addresses:");
            for (int i = 0; i < points.Length; i++)
            {
                float x = points[i][0];
                float y = points[i][1];

                // Convert the point coordinates (in points) to column/row indices.
                // Use Aspose.Cells helper methods to get pixel dimensions of rows/columns.
                int column = (int)Math.Floor(x / sheet.Cells.GetColumnWidthPixel(0));
                int row = (int)Math.Floor(y / sheet.Cells.GetRowHeightPixel(0));

                // Clamp indices to valid worksheet range
                column = Math.Max(0, Math.Min(column, sheet.Cells.MaxColumn));
                row = Math.Max(0, Math.Min(row, sheet.Cells.MaxRow));

                // Convert row/column indices to an Excel cell name (e.g., "B3")
                string cellName = CellsHelper.CellIndexToName(row, column);

                Console.WriteLine($"Point {i + 1}: X={x}, Y={y} => Cell {cellName}");
            }

            // Save the workbook (optional)
            string outputPath = "ShapeConnectionPoints.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
