// Title: Change a Shape's Linked Cell from A1 to B2 with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a ListBox shape, link it to cell A1, reassign the linked cell to B2 using SetLinkedCell, verify the change via the LinkedCell property, update the shape with UpdateSelectedValue, and optionally save the file.
// Keywords: Aspose.Cells | C# | .NET | shape linked cell | SetLinkedCell | ListBox shape | UpdateSelectedValue | programmatic cell link | Excel shape example
// Common Searches: Aspose.Cells change linked cell of shape | SetLinkedCell C# example | verify shape linked cell update Aspose | update ListBox linked cell programmatically | Aspose.Cells shape linked cell verification
// Developer Intent: Reassign a shape's linked cell from A1 to B2 and confirm the shape reads the new cell.
// Use Cases: Switch a form control's data source to a different cell at runtime. | Refresh a shape's displayed value after modifying its linked cell. | Persist linked‑cell changes by saving the workbook.
// AI Prompts: Write C# code that uses Aspose.Cells to change a shape's linked cell from $A$1 to $B$2 and prints the LinkedCell before and after. | Show how to call UpdateSelectedValue after changing a shape's linked cell in Aspose.Cells for .NET. | Provide a complete example that verifies the linked cell update and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a ListBox shape, link it to cell A1, reassign the linked cell to B2 using SetLinkedCell, verify the change via the LinkedCell property, update the shape with UpdateSelectedValue, and optionally save the file.
class ShapeLinkedCellChangeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a ListBox shape (any shape that supports linked cells)
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape listBoxShape = sheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

        // Initially link the shape to cell A1
        listBoxShape.SetLinkedCell("$A$1", false, true);

        // Verify initial linked cell
        Console.WriteLine("Initial LinkedCell: " + listBoxShape.LinkedCell); // Expected: $A$1

        // Change the linked cell from A1 to B2
        listBoxShape.SetLinkedCell("$B$2", false, true);

        // Verify that the linked cell has been updated
        Console.WriteLine("Updated LinkedCell: " + listBoxShape.LinkedCell); // Expected: $B$2

        // Set a value in the new linked cell to demonstrate that the shape reads it
        sheet.Cells["B2"].Value = true; // For a ListBox this would affect selection if input range matches

        // Update the shape's selected value based on the linked cell
        listBoxShape.UpdateSelectedValue();

        // (Optional) If the shape were a CheckBox, you could cast and check IsChecked, but ListBox does not expose a direct value.
        // The successful call to UpdateSelectedValue indicates the shape has read the new linked cell.

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ShapeLinkedCellChangeDemo.xlsx");
    }
}
