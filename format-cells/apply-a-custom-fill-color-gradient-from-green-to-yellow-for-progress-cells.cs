using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ProgressCellGradient
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample progress values (0-100)
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Progress");
        sheet.Cells["A2"].PutValue("Task 1");
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["A3"].PutValue("Task 2");
        sheet.Cells["B3"].PutValue(45);
        sheet.Cells["A4"].PutValue("Task 3");
        sheet.Cells["B4"].PutValue(70);
        sheet.Cells["A5"].PutValue("Task 4");
        sheet.Cells["B5"].PutValue(95);

        // Define the range that will display the gradient (progress column)
        CellArea progressArea = new CellArea
        {
            StartRow = 1,
            EndRow = 4,
            StartColumn = 1,
            EndColumn = 1
        };

        // Create a style with a two‑color gradient from green to yellow
        Style gradientStyle = workbook.CreateStyle();
        gradientStyle.SetTwoColorGradient(
            Color.Green,          // start color
            Color.Yellow,         // end color
            GradientStyleType.Vertical, // gradient direction
            1);                   // variant (1‑4)

        // Apply the gradient style to each cell in the progress range
        for (int row = progressArea.StartRow; row <= progressArea.EndRow; row++)
        {
            Cell cell = sheet.Cells[row, progressArea.StartColumn];
            cell.SetStyle(gradientStyle);
        }

        // Save the workbook
        workbook.Save("ProgressGradient.xlsx");
    }
}