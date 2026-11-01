// Title: How to reassign a rectangle shape’s linked cell from A1 to B2 and refresh its text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that changes a shape’s AlternativeText from "A1" to "B2", updates the shape’s Text property with the new cell value, and prints verification messages using Aspose.Cells. | Demonstrate how to rebind an existing rectangle shape to a different worksheet cell and synchronize its displayed text in a .NET workbook.
// Common Searches: aspocells c# reassign shape linked cell A1 to B2 | update rectangle shape text after moving linked cell in Aspose.Cells | change shape reference to a different worksheet cell in Aspose.Cells | programmatically link Excel shape to a new cell using Aspose.Cells .NET | verify shape text reflects new linked cell value Aspose.Cells
// Tags: Aspose.Cells shape linked cell modification | C# modify shape cell reference | Aspose.Cells rectangle shape text synchronization | programmatic Excel shape cell binding | shape text refresh after cell change Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape, links it to cell A1 via AlternativeText, sets the shape text from A1, then changes the link to B2, updates the shape text, and prints verification messages to confirm the update.
class ShapeLinkedCellDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            // Parameters: shape type, upper left row, upper left column,
            // upper left row offset, upper left column offset,
            // lower right row, lower right column
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // upper left row offset
                0,   // upper left column offset
                4,   // lower right row
                4    // lower right column
            );

            // Set explicit size for the shape
            shape.Width = 120;
            shape.Height = 40;

            // Initially link the shape to cell A1 (using AlternativeText to store the reference)
            shape.AlternativeText = "A1";
            // Set a value in A1; the shape will display this value
            sheet.Cells["A1"].PutValue("First Value");
            // Update shape text to reflect linked cell value
            shape.Text = sheet.Cells["A1"].StringValue;

            // Verify initial linking
            Console.WriteLine($"Initial linked cell: {shape.AlternativeText}");
            Console.WriteLine($"Shape text (should be 'First Value'): {shape.Text}");

            // Change the linked cell to B2
            shape.AlternativeText = "B2";
            // Set a value in B2; the shape will now display this new value
            sheet.Cells["B2"].PutValue("Second Value");
            // Update shape text to reflect new linked cell value
            shape.Text = sheet.Cells["B2"].StringValue;

            // Verify updated linking
            Console.WriteLine($"Updated linked cell: {shape.AlternativeText}");
            Console.WriteLine($"Shape text (should be 'Second Value'): {shape.Text}");

            // Optional: Save the workbook to verify the changes visually
            // workbook.Save("ShapeLinkedCellDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
