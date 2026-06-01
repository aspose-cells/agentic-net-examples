using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddTextboxExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the textbox position (row, column) and size (height, width) in pixels
        int topRow = 2;      // Upper‑left row index (0‑based)
        int leftColumn = 1;  // Upper‑left column index (0‑based)
        int height = 100;    // Height of the textbox in pixels
        int width = 200;     // Width of the textbox in pixels

        // Add a textbox to the worksheet using the TextBoxCollection.Add method
        int textboxIndex = worksheet.TextBoxes.Add(topRow, leftColumn, height, width);

        // Retrieve the added textbox and set its content
        TextBox textbox = worksheet.TextBoxes[textboxIndex];
        textbox.Text = "Hello, Aspose!";

        // Save the workbook to a file
        workbook.Save("AddTextbox.xlsx");
    }
}