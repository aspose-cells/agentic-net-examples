// Title: Add a CheckBox to an Excel worksheet and link it to cell B2 with Aspose.Cells for .NET
// Description: C# example that creates a workbook, inserts a CheckBox at B2, sets its caption and default state, links the control to cell B2 using the LinkedCell property, and saves the file as CheckBoxLinked.xlsx.
// Keywords: Aspose.Cells | .NET | C# | CheckBox control | LinkedCell property | Excel form control | add checkbox to worksheet | cell B2 link | CheckBoxLinked.xlsx | AddCheckBoxLinkedToCell
// Common Searches: Aspose.Cells add checkbox C# | link checkbox to cell B2 Aspose.Cells | set LinkedCell property in Aspose.Cells | create Excel form controls with Aspose.Cells .NET | save workbook with linked checkbox
// Developer Intent: Insert a CheckBox shape into a worksheet and bind its checked state to cell B2.
// Use Cases: Provide a terms‑acceptance box that updates B2, enabling formulas to react to user consent. | Drive conditional formatting or calculations by toggling a linked checkbox stored in B2. | Generate a reusable template where the checkbox state is persisted in a cell for downstream processing.
// AI Prompts: Generate code to add several CheckBox controls, each linked to a different cell, using Aspose.Cells for .NET. | Explain how to read the value of a linked cell after a user changes the checkbox and how to update the checkbox programmatically. | Show how to change a CheckBox caption and size dynamically based on other worksheet cell values.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, inserts a CheckBox at B2, sets its caption and default state, links the control to cell B2 using the LinkedCell property, and saves the file as CheckBoxLinked.xlsx.
class AddCheckBoxLinkedToCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a CheckBox to the worksheet.
        // Parameters: upper left row index, upper left column index, height (pixels), width (pixels)
        int checkBoxIndex = sheet.CheckBoxes.Add(1, 1, 20, 100); // B2 position (row 1, column 1)
        CheckBox checkBox = sheet.CheckBoxes[checkBoxIndex];

        // Set optional properties
        checkBox.Text = "Accept Terms";
        checkBox.Value = true;               // Checked by default
        checkBox.LinkedCell = "B2";           // Link the checkbox state to cell B2

        // Save the workbook
        workbook.Save("CheckBoxLinked.xlsx");
    }
}
