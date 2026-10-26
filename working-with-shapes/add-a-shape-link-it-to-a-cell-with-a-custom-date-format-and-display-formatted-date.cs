// Title: Create a rectangle shape linked to a formatted date cell and display the date using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link shape text to a cell with a custom date format | How to show a formatted date inside a rectangle shape with Aspose.Cells | C# Aspose.Cells set shape placement to move and size with cells | Add rectangle shape to worksheet and bind it to cell A1 date value Aspose.Cells | Apply custom date format to a cell and reference it from a shape in Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | link shape text to cell formula Aspose.Cells | custom date format cell Aspose.Cells | shape placement moveandsize Aspose.Cells | save workbook with linked shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes the current date to cell A1, applies a custom "dd-mmm-yyyy" format, adds a rectangle shape at row 5 column 5, links the shape's text to the formatted date using the =A1 formula, sets the shape to move and size with the cells, and saves the file as ShapeLinkedDate.xlsx.
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

            // Put a date value into cell A1
            Cell dateCell = sheet.Cells["A1"];
            dateCell.PutValue(DateTime.Now);

            // Apply a custom date format to the cell
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-mmm-yyyy";
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                5,   // upper left column
                100, // top offset (in points)
                30,  // left offset (in points)
                30,  // height (in points)
                100  // width (in points)
            );

            // Link the shape's text to cell A1 so it displays the formatted date
            shape.Text = "=A1";

            // Ensure the shape moves and sizes with the cells (optional)
            shape.Placement = PlacementType.MoveAndSize;

            // Save the workbook to a file
            workbook.Save("ShapeLinkedDate.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
