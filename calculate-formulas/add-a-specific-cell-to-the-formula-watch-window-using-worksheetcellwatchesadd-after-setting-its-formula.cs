using System;
using Aspose.Cells;

namespace AsposeCellsWatchDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell A1 (e.g., sum of B1 and C1)
            sheet.Cells["A1"].Formula = "=SUM(B1,C1)";

            // Optionally put some values in B1 and C1 for the formula to calculate
            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["C1"].PutValue(20);

            // Add the cell to the Formula Watch Window using its address
            int watchIndex = sheet.CellWatches.Add("A1");

            // Retrieve the added CellWatch (optional, just to demonstrate access)
            CellWatch watch = sheet.CellWatches[watchIndex];
            Console.WriteLine($"Added watch for cell: {watch.CellName} (Row={watch.Row}, Column={watch.Column})");

            // Save the workbook to a file
            workbook.Save("FormulaWatchDemo.xlsx");
        }
    }
}