using System;
using Aspose.Cells;

namespace AsposeCellsComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate input cells
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Set a formula in B1 that depends on A1 and A2
            cells["B1"].Formula = "=A1+A2";

            // ---------- Calculate using Workbook.CalculateFormula ----------
            workbook.CalculateFormula(); // calculates all formulas in the workbook
            object workbookResult = cells["B1"].Value; // store the result after workbook calculation

            // ---------- Calculate using Cell.Calculate ----------
            // Recalculate the same cell individually
            cells["B1"].Calculate(new CalculationOptions());
            object cellResult = cells["B1"].Value; // store the result after cell calculation

            // Compare the two results for consistency
            bool areEqual = Equals(workbookResult, cellResult);

            // Output the comparison result
            Console.WriteLine($"Result from Workbook.CalculateFormula: {workbookResult}");
            Console.WriteLine($"Result from Cell.Calculate: {cellResult}");
            Console.WriteLine($"Are both results equal? {areEqual}");
        }
    }
}