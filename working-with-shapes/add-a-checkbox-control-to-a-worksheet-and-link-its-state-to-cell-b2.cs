using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddCheckBoxLinkedToCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a CheckBox to the worksheet
        // Parameters: upper left row, upper left column, height (pixels), width (pixels)
        int checkBoxIndex = sheet.CheckBoxes.Add(5, 5, 20, 100);
        CheckBox checkBox = sheet.CheckBoxes[checkBoxIndex];

        // Set the display text of the CheckBox
        checkBox.Text = "Accept Terms";

        // Link the CheckBox state to cell B2
        checkBox.LinkedCell = "B2";

        // Optionally set the initial checked state
        checkBox.Value = true;

        // Save the workbook to a file
        workbook.Save("CheckBoxLinked.xlsx");
    }
}