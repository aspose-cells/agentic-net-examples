using System;
using Aspose.Cells;

namespace AutoFitAndFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Short");
            sheet.Cells["B1"].PutValue("This is a longer text that will require column auto‑fit");
            sheet.Cells["C1"].PutValue("Medium length");
            sheet.Cells["A2"].PutValue("Another row with a very very long piece of text to test column width adjustment");
            sheet.Cells["B2"].PutValue(12345);
            sheet.Cells["C2"].PutValue(DateTime.Now);

            // Auto‑fit all columns so their widths match the content
            sheet.AutoFitColumns();

            // Freeze panes at cell C3 (row index 2, column index 2) with 2 rows and 2 columns frozen
            // This ensures the frozen area uses the column widths set by AutoFitColumns
            sheet.FreezePanes(2, 2, 2, 2);

            // Save the workbook
            workbook.Save("AutoFitAndFreeze.xlsx");
        }
    }
}