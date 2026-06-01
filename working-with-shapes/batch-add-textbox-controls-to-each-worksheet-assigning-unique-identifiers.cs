using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class BatchAddTextBoxes
{
    static void Main()
    {
        // Create a new workbook (default contains one worksheet)
        Workbook workbook = new Workbook();

        // Optionally add more worksheets for demonstration
        workbook.Worksheets.Add();
        workbook.Worksheets.Add();

        // Iterate through each worksheet in the workbook
        for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
        {
            Worksheet sheet = workbook.Worksheets[sheetIdx];

            // Add a TextBox to the worksheet.
            // Parameters: topRow, leftColumn, height (pixels), width (pixels)
            int textBoxIndex = sheet.TextBoxes.Add(1, 1, 100, 200);

            // Retrieve the added TextBox object
            TextBox textBox = sheet.TextBoxes[textBoxIndex];

            // Assign a unique name/identifier to the TextBox
            textBox.Name = $"TextBox_Sheet{sheetIdx}_Idx{textBoxIndex}";

            // Optionally set some display text
            textBox.Text = $"TextBox on Sheet {sheetIdx}";
        }

        // Save the workbook with the added TextBoxes
        workbook.Save("BatchTextBoxes.xlsx");
    }
}