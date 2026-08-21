// Title: Monitor a Single Cell and Trigger Formula Recalculation with Aspose.Cells for .NET
// Description: Demonstrates how to set Aspose.Cells to manual calculation mode and attach a custom AbstractCalculationMonitor that fires only when a designated cell changes. The example logs the original and new values, recalculates dependent formulas, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | AbstractCalculationMonitor | custom calculation monitor | manual calculation mode | cell change event | formula recalculation trigger | monitor specific cell | Workbook events
// Common Searches: Aspose.Cells monitor single cell change | How to use AbstractCalculationMonitor in C# | Manual calculation mode Aspose.Cells example | Trigger formula recalculation only for changed cell | Detect cell value change with Aspose.Cells
// Developer Intent: Enable workbook events that recalculate formulas only when a designated cell is modified, using a custom calculation monitor in Aspose.Cells for .NET.
// Use Cases: Log modifications to a key parameter cell and update dependent results only when that cell changes. | Reduce processing time in large workbooks by performing manual calculations and monitoring critical cells. | Create an audit trail of specific cell edits during manual recalculation.
// AI Prompts: Generate C# code that uses Aspose.Cells to monitor cell D5 with a custom AbstractCalculationMonitor and log its original and new values. | Explain how manual calculation mode combined with a calculation monitor can prevent unnecessary formula evaluations in a workbook containing thousands of formulas. | Provide step‑by‑step guidance to extend the example so it monitors multiple cells (e.g., B2, C3) simultaneously.

using System;
using Aspose.Cells;

namespace WorkbookEventDemo
{
    // Custom calculation monitor to react only when a specific cell is recalculated
    // Demonstrates how to set Aspose.Cells to manual calculation mode and attach a custom AbstractCalculationMonitor that fires only when a designated cell changes. The example logs the original and new values, recalculates dependent formulas, and saves the workbook.
    class SpecificCellCalculationMonitor : AbstractCalculationMonitor
    {
        // Define the cell we are interested in (row and column indexes, zero‑based)
        private readonly int _targetRow;
        private readonly int _targetColumn;

        public SpecificCellCalculationMonitor(int targetRow, int targetColumn)
        {
            _targetRow = targetRow;
            _targetColumn = targetColumn;
        }

        // Called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Check if the calculated cell matches the target cell
            if (rowIndex == _targetRow && colIndex == _targetColumn)
            {
                // ValueChanged indicates whether the cell value actually changed during calculation
                if (ValueChanged)
                {
                    Console.WriteLine($"Target cell changed: Sheet{sheetIndex} " +
                                      $"R{rowIndex + 1}C{colIndex + 1} " +
                                      $"from [{OriginalValue}] to [{CalculatedValue}]");
                }
                else
                {
                    Console.WriteLine($"Target cell recalculated but value unchanged: " +
                                      $"Sheet{sheetIndex} R{rowIndex + 1}C{colIndex + 1}");
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create ----------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet

            // Set up sample data and formulas
            sheet.Cells["A1"].PutValue(5);                         // A1 = 5
            sheet.Cells["B1"].PutValue(10);                        // B1 = 10 (this is the cell we will monitor)
            sheet.Cells["C1"].Formula = "=A1+B1";                  // C1 = A1 + B1

            // ---------- Configure calculation ----------
            // Use manual calculation mode so formulas are not auto‑recalculated
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Create calculation options with our custom monitor for cell B1 (row 0, column 1)
            CalculationOptions options = new CalculationOptions();
            options.CalculationMonitor = new SpecificCellCalculationMonitor(0, 1);

            // Initial calculation (optional, just to show baseline)
            workbook.CalculateFormula(options);
            Console.WriteLine($"Initial C1 value: {sheet.Cells["C1"].Value}");

            // ---------- Modify a specific cell ----------
            // Change B1 – this should trigger the monitor during the next calculation
            sheet.Cells["B1"].PutValue(20);
            Console.WriteLine("B1 modified to 20.");

            // ---------- Trigger recalculation ----------
            // Because we are in manual mode, we explicitly call CalculateFormula.
            // The monitor will fire only for the target cell (B1).
            workbook.CalculateFormula(options);
            Console.WriteLine($"After change C1 value: {sheet.Cells["C1"].Value}");

            // ---------- Save ----------
            workbook.Save("WorkbookEventDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
