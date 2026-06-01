using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a button at cell D5 (row index 4, column index 3)
        // Parameters: topRow, top offset, leftColumn, left offset, height, width
        Button button = sheet.Shapes.AddButton(4, 0, 3, 0, 30, 100);
        button.Text = "Run Macro";
        button.MacroName = "MyMacro";

        // Save the workbook
        workbook.Save("ButtonWithMacro.xlsx");
    }
}