using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a formula in cell B2 (e.g., sum of A1:A5)
        worksheet.Cells["B2"].Formula = "=SUM(A1:A5)";

        // Add a TextBox shape to the worksheet
        // Parameters: upperRow, leftColumn, upperRowOffset, leftColumnOffset, width, height
        TextBox textBox = worksheet.Shapes.AddTextBox(2, 2, 0, 0, 150, 30);

        // Link the TextBox text to the cell containing the formula
        // false = A1 style, true = locale aware
        textBox.SetLinkedCell("$B$2", false, true);

        // Save the workbook to a file
        workbook.Save("LinkedShapeDemo.xlsx");
    }
}