using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsCalculationChainDemo
{
    class Program
    {
        static void Main()
        {
            // Create a workbook and add a large number of formulas
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with values 1..10000
            int rowCount = 10000;
            for (int i = 0; i < rowCount; i++)
            {
                cells[i, 0].PutValue(i + 1);
            }

            // In column B place a formula that sums the range A1:A{row}
            for (int i = 0; i < rowCount; i++)
            {
                cells[i, 1].Formula = $"=SUM(A1:A{i + 1})";
            }

            // ------------------- Calculation without chain -------------------
            wb.Settings.FormulaSettings.EnableCalculationChain = false;
            Stopwatch sw = Stopwatch.StartNew();
            wb.CalculateFormula();               // full calculation
            sw.Stop();
            Console.WriteLine($"Calculation time without chain: {sw.ElapsedMilliseconds} ms");

            // ------------------- Calculation with chain -------------------
            // Create a fresh workbook with the same data to avoid cached results
            Workbook wbChain = new Workbook();
            Worksheet sheetChain = wbChain.Worksheets[0];
            Cells cellsChain = sheetChain.Cells;

            for (int i = 0; i < rowCount; i++)
            {
                cellsChain[i, 0].PutValue(i + 1);
            }
            for (int i = 0; i < rowCount; i++)
            {
                cellsChain[i, 1].Formula = $"=SUM(A1:A{i + 1})";
            }

            wbChain.Settings.FormulaSettings.EnableCalculationChain = true;
            Stopwatch swChain = Stopwatch.StartNew();
            wbChain.CalculateFormula();          // first calculation builds the chain
            swChain.Stop();
            Console.WriteLine($"First calculation with chain (build time): {swChain.ElapsedMilliseconds} ms");

            // Modify a small part and recalculate to see the benefit
            cellsChain[0, 0].PutValue(9999);      // change only the first value
            Stopwatch swRecalc = Stopwatch.StartNew();
            wbChain.CalculateFormula();          // subsequent calculation uses the chain
            swRecalc.Stop();
            Console.WriteLine($"Recalculation after small change with chain: {swRecalc.ElapsedMilliseconds} ms");

            // Save the workbooks (optional, demonstrates lifecycle usage)
            wb.Save("WithoutChain.xlsx", SaveFormat.Xlsx);
            wbChain.Save("WithChain.xlsx", SaveFormat.Xlsx);
        }
    }
}