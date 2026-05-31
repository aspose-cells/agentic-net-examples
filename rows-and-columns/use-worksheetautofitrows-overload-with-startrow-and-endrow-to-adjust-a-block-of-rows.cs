using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitRowsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some rows with varying text lengths to demonstrate autofit
            worksheet.Cells["A1"].PutValue("Short text");
            worksheet.Cells["A2"].PutValue("This is a longer piece of text that should cause the row height to increase when autofit is applied.");
            worksheet.Cells["A3"].PutValue("Another short text");
            worksheet.Cells["A4"].PutValue("A very long text that spans multiple lines when wrapped. It will be used to test the AutoFitRows method over a specific range of rows.");
            worksheet.Cells["A5"].PutValue("Final short text");

            // Enable text wrapping for the cells that contain long text
            Style wrapStyle = worksheet.Cells["A2"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            worksheet.Cells["A2"].SetStyle(wrapStyle);
            worksheet.Cells["A4"].SetStyle(wrapStyle);

            // AutoFit rows from index 1 to 3 (i.e., rows 2 to 4)
            // This uses the Worksheet.AutoFitRows(int startRow, int endRow) overload
            worksheet.AutoFitRows(1, 3);

            // Save the workbook (lifecycle save)
            workbook.Save("AutoFitRowsBlockDemo.xlsx");
        }
    }
}