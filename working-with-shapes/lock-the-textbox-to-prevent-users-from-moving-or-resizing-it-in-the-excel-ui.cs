// Title: How to lock a TextBox shape in an Excel worksheet to prevent moving or resizing using Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook, insert a TextBox at a given cell, set IsLocked = true, and save the file with Aspose.Cells in C#. | Generate an Excel file where the TextBox shape is locked and set to FreeFloating so users cannot adjust its size or position, using the Aspose.Cells .NET API. | Write C# code that adds a TextBox to a worksheet, disables moving and resizing by locking the shape, and keeps it independent of cell changes with Aspose.Cells.
// Common Searches: Aspose.Cells C# example to prevent textbox movement in Excel UI | how to disable resizing of a shape in Excel using Aspose.Cells .NET | set shape IsLocked property in Aspose.Cells workbook C# | make textbox shape non-editable in Excel with Aspose.Cells | Aspose.Cells lock shape placement FreeFloating C# tutorial
// Tags: Aspose.Cells IsLocked property for shapes | C# Aspose.Cells textbox shape protection | FreeFloating placement for Aspose.Cells shapes | Excel shape immovable using Aspose.Cells | Aspose.Cells workbook shape lock example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a TextBox shape at a specified location, assigns text, locks the shape by setting IsLocked to true, optionally sets its placement to FreeFloating, and saves the workbook as LockedTextbox.xlsx.
class LockTextboxExample
{
    static void Main()
    {
        // Load an existing workbook or create a new one
        Workbook workbook = new Workbook(); // creates a new workbook
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        int upperLeftRow = 2;
        int upperLeftColumn = 2;
        int top = 5;
        int left = 5;
        int height = 50;
        int width = 150;

        // Create the textbox shape
        TextBox textbox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, top, left, height, width);
        textbox.Text = "Locked TextBox";

        // Lock the textbox to prevent moving or resizing in the Excel UI
        textbox.IsLocked = true;

        // Optionally, set the shape's placement to FreeFloating (default) to keep it independent of cells
        textbox.Placement = PlacementType.FreeFloating;

        // Save the workbook to a file
        workbook.Save("LockedTextbox.xlsx");
    }
}
