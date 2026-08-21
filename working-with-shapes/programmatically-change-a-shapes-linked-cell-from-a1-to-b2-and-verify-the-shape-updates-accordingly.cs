// Title: Change a Shape’s Linked Cell from A1 to B2 using Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds a rectangle shape, links it to cell A1, reassigns the link to B2, updates the displayed value, and saves the file so you can confirm the shape reflects the new cell.
// Keywords: Aspose.Cells | C# | shape linked cell | SetLinkedCell | UpdateSelectedValue | rectangle shape | Excel automation | programmatic linked cell change | Aspose.Cells for .NET | cell reference
// Common Searches: Aspose.Cells change shape linked cell | SetLinkedCell C# example | Update shape after linked cell change Aspose | How to reassign linked cell of a shape in Excel using Aspose.Cells | Verify shape linked cell update programmatically
// Developer Intent: Reassign a shape’s linked cell and validate the update.
// Use Cases: Generate reports where shapes need to point to different data cells dynamically | Build Excel templates that adjust form control links during runtime | Create interactive dashboards that switch shape references based on user input
// AI Prompts: Provide C# code that uses Aspose.Cells to move a shape’s linked cell from $A$1 to $B$2 and refresh its value. | Explain the role of SetLinkedCell and UpdateSelectedValue when modifying shape references in a workbook. | Write a unit test in C# that asserts the LinkedCell property changes from $A$1 to $B$2 after calling SetLinkedCell.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds a rectangle shape, links it to cell A1, reassigns the link to B2, updates the displayed value, and saves the file so you can confirm the shape reflects the new cell.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, height, width, upper left row offset, upper left column offset
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

        // Initially link the shape to cell A1
        shape.SetLinkedCell("$A$1", false, true);

        // Output the initial linked cell
        Console.WriteLine("Initial LinkedCell: " + shape.LinkedCell);

        // Change the linked cell from A1 to B2
        shape.SetLinkedCell("$B$2", false, true);

        // Output the updated linked cell to verify the change
        Console.WriteLine("Updated LinkedCell: " + shape.LinkedCell);

        // Optionally, set a value in the new linked cell and update the shape's selected value
        worksheet.Cells["B2"].Value = "Sample";
        shape.UpdateSelectedValue();

        // Save the workbook (verification can be done by opening the file)
        workbook.Save("LinkedCellChangeDemo.xlsx");
    }
}
