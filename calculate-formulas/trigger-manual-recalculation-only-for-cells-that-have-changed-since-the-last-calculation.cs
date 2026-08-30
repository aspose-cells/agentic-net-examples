// Title: How to trigger manual recalculation only for cells changed since the last calculation using Aspose.Cells C#
// AI Prompts: Generate C# code that sets an Aspose.Cells workbook to manual calculation mode, modifies a source cell, and calls CalculateFormula with a custom calculation monitor to capture changed cells. | Show how to implement a subclass of AbstractCalculationMonitor that records the address of any cell whose value changes during a manual calculation in Aspose.Cells. | Provide a step‑by‑step example that resets the monitor, updates a cell, runs workbook.CalculateFormula, and prints the dependent cell that was recalculated.
// Common Searches: Aspose.Cells C# manual calculation mode recalculate only changed cells | C# detect which cells changed after workbook.CalculateFormula in Aspose.Cells | How to use AbstractCalculationMonitor to log changed cells in Aspose.Cells | Prevent full workbook recalculation in Aspose.Cells by tracking modified cells | Example of manual formula evaluation with dependent cell detection in Aspose.Cells .NET
// Tags: Aspose.Cells calculation settings | cell change detection Aspose.Cells | selective formula evaluation Aspose.Cells | user-defined calculation monitor Aspose.Cells | track dependent cells Aspose.Cells

using System;
using Aspose.Cells;
using System.Collections;

namespace AsposeCellsManualRecalcDemo
{
    // Custom monitor to detect which cells changed after a calculation
    // The example creates a workbook, switches its calculation mode to Manual, defines a ChangeDetectionMonitor derived from AbstractCalculationMonitor to capture the address of any cell whose value changes during CalculateFormula, performs an initial calculation, modifies source cells, triggers manual recalculation while logging changed cells, and finally saves the workbook.
    class ChangeDetectionMonitor : AbstractCalculationMonitor
    {
        // Stores the last changed cell address (for demo purposes)
        public string LastChangedCellAddress { get; private set; } = string.Empty;

        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // If the cell value was changed, remember its address
            if (ValueChanged)
            {
                // Convert zero‑based indexes to A1 style
                string address = CellsHelper.CellIndexToName(rowIndex, colIndex);
                LastChangedCellAddress = address;
                Console.WriteLine($"Cell {address} changed from [{OriginalValue}] to [{CalculatedValue}]");
            }
        }

        // Reset the stored address before the next calculation pass
        public void Reset()
        {
            LastChangedCellAddress = string.Empty;
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (lifecycle rule)
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 2. Put initial data and formulas
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Source values
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Formula that depends on A1 and A2
            cells["B1"].Formula = "=A1+A2";

            // -------------------------------------------------
            // 3. Switch calculation mode to Manual
            // -------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Optional: prevent automatic calculation on save
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // -------------------------------------------------
            // 4. Prepare calculation options with a custom monitor
            // -------------------------------------------------
            ChangeDetectionMonitor monitor = new ChangeDetectionMonitor();
            CalculationOptions calcOptions = new CalculationOptions
            {
                CalculationMonitor = monitor,
                // We do not need recursive recalculation of unrelated cells
                Recursive = true
            };

            // -------------------------------------------------
            // 5. First manual calculation (establish baseline)
            // -------------------------------------------------
            Console.WriteLine("=== First manual calculation ===");
            workbook.CalculateFormula(calcOptions);
            Console.WriteLine($"B1 value after first calc: {cells["B1"].Value}");
            monitor.Reset();

            // -------------------------------------------------
            // 6. Change a source cell – only this change should trigger recalculation
            // -------------------------------------------------
            Console.WriteLine("\n=== Changing A1 ===");
            cells["A1"].PutValue(30);   // modify source

            // Manual trigger – only changed cells (and their dependents) are recalculated
            workbook.CalculateFormula(calcOptions);

            // After calculation the monitor tells us which cell actually changed
            if (!string.IsNullOrEmpty(monitor.LastChangedCellAddress))
            {
                Console.WriteLine($"Recalculated cell: {monitor.LastChangedCellAddress}");
                Console.WriteLine($"New B1 value: {cells["B1"].Value}");
            }
            else
            {
                Console.WriteLine("No cell value changed.");
            }

            monitor.Reset();

            // -------------------------------------------------
            // 7. Change a cell that does NOT affect any formula
            // -------------------------------------------------
            Console.WriteLine("\n=== Changing C1 (no dependent formulas) ===");
            cells["C1"].PutValue(999); // independent cell

            // Manual trigger – because C1 has no dependents, nothing should be recalculated
            workbook.CalculateFormula(calcOptions);

            if (string.IsNullOrEmpty(monitor.LastChangedCellAddress))
                Console.WriteLine("No dependent cell changed, as expected.");

            // -------------------------------------------------
            // 8. Save the workbook (lifecycle rule)
            // -------------------------------------------------
            workbook.Save("ManualRecalcDemo.xlsx");
            Console.WriteLine("\nWorkbook saved as ManualRecalcDemo.xlsx");
        }
    }
}
