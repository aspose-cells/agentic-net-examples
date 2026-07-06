using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaComparison
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and set up sample data and a formula.
            // ------------------------------------------------------------
            Workbook wbOriginal = new Workbook();
            Worksheet wsOriginal = wbOriginal.Worksheets[0];

            wsOriginal.Cells["A1"].PutValue(5);               // First operand
            wsOriginal.Cells["B1"].PutValue(10);              // Second operand
            wsOriginal.Cells["C1"].Formula = "=A1+B1";        // Formula to be tested

            // ------------------------------------------------------------
            // 2. Clone the workbook into a second instance.
            //    The clone will be used for full workbook calculation.
            // ------------------------------------------------------------
            MemoryStream ms = new MemoryStream();
            wbOriginal.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0;
            Workbook wbClone = new Workbook(ms);
            Worksheet wsClone = wbClone.Worksheets[0];

            // ------------------------------------------------------------
            // 3. Calculate only the target cell using Cell.Calculate().
            // ------------------------------------------------------------
            Cell targetCellOriginal = wsOriginal.Cells["C1"];
            targetCellOriginal.Calculate(new CalculationOptions());
            object resultCellCalculate = targetCellOriginal.Value;

            // ------------------------------------------------------------
            // 4. Calculate the entire workbook using Workbook.CalculateFormula().
            // ------------------------------------------------------------
            wbClone.CalculateFormula();
            object resultWorkbookCalculate = wsClone.Cells["C1"].Value;

            // ------------------------------------------------------------
            // 5. Compare the two results for consistency.
            // ------------------------------------------------------------
            bool areEqual = Equals(resultCellCalculate, resultWorkbookCalculate);

            Console.WriteLine($"Result from Cell.Calculate():      {resultCellCalculate}");
            Console.WriteLine($"Result from Workbook.CalculateFormula(): {resultWorkbookCalculate}");
            Console.WriteLine($"Are both results equal? {areEqual}");
        }
    }
}