// Title: Unlock a locked TextBox shape in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a TextBox shape that is initially locked, enable editing of drawing objects on a protected worksheet, set the shape's IsLocked property to false, optionally protect the sheet, and save the file so the TextBox can be moved and resized.
// Keywords: Aspose.Cells | .NET | C# | TextBox shape | IsLocked | Worksheet protection | AllowEditingObject | unlock shape | resize shape | move shape
// Common Searches: Aspose.Cells unlock TextBox | How to edit size of a locked shape in .NET | Enable shape movement on a protected worksheet | C# unlock textbox shape Aspose.Cells | Allow editing of drawing objects after sheet protection
// Developer Intent: Remove the lock from a TextBox shape so its position and dimensions can be modified programmatically.
// Use Cases: Programmatically release a TextBox to reposition it on a protected sheet. | Allow end‑users to resize or move shapes while the worksheet remains locked for data entry. | Adjust the dimensions of a shape after unlocking it to accommodate dynamic content.
// AI Prompts: Generate C# code with Aspose.Cells that unlocks a TextBox shape and then changes its width and height. | Show how to protect a worksheet but still permit editing of drawing objects, including unlocking a shape. | Explain the role of the AllowEditingObject property when unlocking shapes on a protected sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UnlockTextBoxDemo
{
    // Demonstrates how to create a workbook, add a TextBox shape that is initially locked, enable editing of drawing objects on a protected worksheet, set the shape's IsLocked property to false, optionally protect the sheet, and save the file so the TextBox can be moved and resized.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a regular TextBox shape (initially locked)
            // Parameters: upper left row, upper left column, top offset, left offset, width, height (in pixels)
            Shape textBoxShape = sheet.Shapes.AddTextBox(2, 2, 0, 0, 150, 60);
            textBoxShape.Text = "Locked TextBox";

            // Lock the shape to simulate a previously locked textbox
            textBoxShape.IsLocked = true;

            // ---- Unlocking process ----

            // 1. Ensure the worksheet allows editing of drawing objects when protected
            sheet.Protection.AllowEditingObject = true;

            // 2. Unlock the textbox shape so its position and size can be edited
            textBoxShape.IsLocked = false;

            // (Optional) Protect the worksheet to see the effect of AllowEditingObject
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("UnlockedTextBox.xlsx");
        }
    }
}
