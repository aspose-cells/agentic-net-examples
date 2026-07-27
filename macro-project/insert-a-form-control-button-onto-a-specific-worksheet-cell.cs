using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertButtonDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a Form Control button anchored to cell B2 (row index 1, column index 1)
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Button button = sheet.Shapes.AddButton(1, 0, 1, 0, 30, 100);
        button.Text = "Press Me";

        // Optionally link the button's value to a cell (e.g., C3)
        button.LinkedCell = "C3";

        // Save the workbook with the button
        workbook.Save("ButtonInCell.xlsx");
    }
}