// Title: How to insert a rectangle shape into a specific cell (e.g., B2) using Aspose.Cells for .NET C#
// Common Searches: C# Aspose.Cells add rectangle shape to cell B2 | Insert Excel shape at row and column coordinates with Aspose.Cells | Aspose.Cells place rectangle at specific worksheet cell using AddShape | Set shape size in points when adding rectangle in Aspose.Cells C# | Create rectangle shape in Excel workbook with Aspose.Cells .NET API
// Tags: AddShape rectangle Aspose.Cells | position shape by row column indices | MsoDrawingType.Rectangle example | set shape height width points | save workbook with inserted shape

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;

// The example creates a new workbook, accesses the first worksheet, and inserts a rectangle shape at cell B2 using the AddShape method with MsoDrawingType.Rectangle. Shape dimensions are defined in points (height 100, width 200). Optional fill and line color customizations are shown, and the workbook is saved as RectangleShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell location (e.g., B2)
            int targetRow = 1;      // Row index (zero‑based) for row 2
            int targetColumn = 1;   // Column index (zero‑based) for column B

            // Define shape size (in points)
            int shapeHeight = 100;
            int shapeWidth = 200;

            // Insert a rectangle shape at the target cell.
            // Use MsoDrawingType.Rectangle for compatibility with older Aspose.Cells versions.
            Shape rectangle = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Rectangle shape type
                targetRow,                // Upper left row
                targetColumn,             // Upper left column
                0,                        // Top offset within the cell
                0,                        // Left offset within the cell
                shapeHeight,              // Height of the shape
                shapeWidth);              // Width of the shape

            // Optional: customize appearance if supported
            // Uncomment the following lines if the version of Aspose.Cells supports these properties.
            // rectangle.FillColor = Color.LightBlue;
            // rectangle.Line.Color = Color.DarkBlue;

            // Save the workbook
            workbook.Save("RectangleShape.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
