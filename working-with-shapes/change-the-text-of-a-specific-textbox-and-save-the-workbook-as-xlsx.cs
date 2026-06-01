using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ChangeTextBoxText
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox to the worksheet (textboxes collection rule)
        // Parameters: upper left row, upper left column, width, height
        int textboxIndex = worksheet.TextBoxes.Add(2, 2, 200, 100);

        // Retrieve the added textbox
        TextBox textbox = worksheet.TextBoxes[textboxIndex];

        // Set the desired text (Text property rule)
        textbox.Text = "New text for the textbox";

        // Save the workbook as XLSX (save method rule)
        workbook.Save("ModifiedTextBox.xlsx", SaveFormat.Xlsx);
    }
}