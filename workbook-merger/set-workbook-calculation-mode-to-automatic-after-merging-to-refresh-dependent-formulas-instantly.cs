using System;
using Aspose.Cells;

namespace AsposeCellsMergeRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Set a formula that sums the range A1:A3
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Merge cells A1:A3 (this may affect the formula)
            sheet.Cells.Merge(0, 0, 3, 1); // Merge rows 0-2, columns 0-0 (A1:A3)

            // Ensure the workbook calculation mode is Automatic
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Recalculate formulas so dependent cells reflect the merge change instantly
            workbook.CalculateFormula();

            // Output the result to console
            Console.WriteLine("Result of B1 after merge and recalculation: " + cells["B1"].Value);

            // Save the workbook
            workbook.Save("MergedRefreshDemo.xlsx");
        }
    }
}