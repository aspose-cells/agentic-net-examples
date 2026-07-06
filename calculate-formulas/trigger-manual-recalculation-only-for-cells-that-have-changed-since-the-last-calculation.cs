using System;
using Aspose.Cells;
using System.Collections;

namespace AsposeCellsExamples
{
    // Custom monitor to report cells whose values changed after calculation
    public class ChangedCellMonitor : AbstractCalculationMonitor
    {
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // ValueChanged is true only when the cell's value differs from the previous value
            if (ValueChanged)
            {
                Console.WriteLine($"Cell {CellReference(rowIndex, colIndex)} changed from [{OriginalValue}] to [{CalculatedValue}]");
            }
        }

        private string CellReference(int row, int col)
        {
            // Convert zero‑based row/col to A1 style reference
            return CellsHelper.CellIndexToName(row, col);
        }
    }

    public class ManualRecalcDemo
    {
        public static void Run()
        {
            // ---------- Create ----------
            Workbook wb = new Workbook();                     // new workbook
            Worksheet ws = wb.Worksheets[0];                  // first worksheet
            Cells cells = ws.Cells;

            // Enable calculation chain so Aspose tracks dependencies
            wb.Settings.FormulaSettings.EnableCalculationChain = true;

            // Populate data and formulas
            cells["A1"].PutValue(10);                         // source value
            cells["A2"].PutValue(20);                         // source value
            cells["B1"].Formula = "=A1*2";                    // depends on A1
            cells["B2"].Formula = "=A2*3";                    // depends on A2
            cells["C1"].Formula = "=B1+B2";                   // depends on B1 and B2

            // Set up calculation monitor
            ChangedCellMonitor monitor = new ChangedCellMonitor();
            CalculationOptions opts = new CalculationOptions
            {
                CalculationMonitor = monitor,
                Recursive = true          // default, keep for clarity
            };

            // First calculation – all formulas are evaluated
            wb.CalculateFormula(opts);
            Console.WriteLine("Initial calculation completed.\n");

            // ---------- Modify ----------
            // Change only A1; only cells that depend on A1 should be recalculated
            cells["A1"].PutValue(15);

            // ---------- Recalculate ----------
            // Because EnableCalculationChain is true, Aspose will recalc only affected cells
            wb.CalculateFormula(opts);
            Console.WriteLine("\nRecalculation after change completed.");

            // ---------- Save ----------
            wb.Save("ManualRecalcDemo.xlsx");
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            ManualRecalcDemo.Run();
        }
    }
}