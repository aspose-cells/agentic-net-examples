// Title: Create a rectangle shape linked to a formatted cell with thousand separators using Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to a worksheet and set its Text property to a formula that references cell A1, which contains a number formatted with "#,##0". | Apply a custom number format "#,##0" to a numeric cell so the value displays with thousand separators. | Configure the shape's Placement to MoveAndSize so it moves and resizes with the underlying cells, then save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells how to bind shape text to a cell value with custom number format | C# add rectangle shape that displays cell A1 formatted with thousand separator | set shape placement move and size in Aspose.Cells .NET example | apply "#,##0" number format to a cell and link it to a shape using Aspose.Cells | save workbook with shape linked to formatted cell Aspose.Cells
// Tags: drawing rectangle object Aspose.Cells | link shape text to cell formula | custom number format "#,##0" Aspose.Cells | placement mode MoveAndSize for shapes .NET | save workbook with linked shape xlsx

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, writes 1234567 to cell A1, applies the custom number format "#,##0" for thousand separators, adds a rectangle shape, links its text to A1, sets the shape to move and size with cells, and saves the file as ShapeLinkedToCell.xlsx.
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

            // Set a numeric value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234567);

            // Apply custom number format with thousand separator
            Style style = cell.GetStyle();
            style.Custom = "#,##0";
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

            // Link the shape's text to cell A1 so it displays the formatted value
            shape.Text = "=A1";

            // Ensure the shape moves and sizes with the cells
            shape.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            string outputPath = "ShapeLinkedToCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
