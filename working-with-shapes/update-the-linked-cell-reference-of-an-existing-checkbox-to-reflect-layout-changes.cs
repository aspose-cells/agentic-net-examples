using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class UpdateCheckBoxLinkedCell
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a checkbox to the worksheet
        // Parameters: upper left row, upper left column, height, width
        int chkIndex = sheet.CheckBoxes.Add(2, 2, 20, 100);
        CheckBox checkBox = sheet.CheckBoxes[chkIndex];

        // Initially link the checkbox to cell B5
        checkBox.LinkedCell = "B5";

        // Simulate a layout change: move the checkbox and change its linked cell to C10
        // (In a real scenario, the movement might affect which cell should be linked)
        checkBox.Top = 150;   // move vertically (pixels)
        checkBox.Left = 200;  // move horizontally (pixels)

        // Update the linked cell reference to reflect the new layout
        checkBox.LinkedCell = "C10";

        // Save the workbook (save rule)
        workbook.Save("UpdatedCheckBoxLinkedCell.xlsx");
    }
}