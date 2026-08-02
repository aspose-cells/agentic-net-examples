// Title: Aspose.Cells for .NET – Manual Recalculation of Only Changed Cells with a Custom CalculationMonitor (C#)
// Description: Shows how to set Aspose.Cells to manual calculation mode, change a source value, and run CalculateFormula with a custom CalculationMonitor that logs only the cells whose values actually changed. Includes an initial full calculation, dependent updates, and saving without automatic recalculation.
// Keywords: Aspose.Cells | C# | .NET | manual calculation mode | CalculationMonitor | changed cells detection | recalculate dependent cells | CalculateFormula | custom monitor | workbook formula evaluation
// Common Searches: Aspose.Cells manual calculation mode C# | how to recalculate only changed cells Aspose.Cells | custom CalculationMonitor example Aspose.Cells | detect changed cell values after formula calculation .NET | skip full workbook recalculation Aspose.Cells
// Developer Intent: The developer needs to perform a manual recalculation that updates and reports only the cells whose values have changed since the previous calculation.
// Use Cases: Log each cell that changed after modifying a source value, showing old and new results. | Reduce processing time by avoiding a full workbook recalculation and updating only dependent cells. | Collect changed cell addresses and values into a collection for change‑report generation before saving the file.
// AI Prompts: Generate C# code that switches Aspose.Cells to manual calculation mode, modifies a source cell, and uses a custom AbstractCalculationMonitor to output only cells with changed values. | Explain step‑by‑step how to configure CalculationOptions with a custom CalculationMonitor to capture changed cell references during a manual recalculation in Aspose.Cells. | Provide a sample that iterates over the cells reported by the monitor, stores their A1 addresses and new values in a list, and writes the list to a CSV file.

using System;
using Aspose.Cells;
using System.Collections;

namespace AsposeCellsRecalcDemo
{
    // Custom monitor to detect cells whose values changed after calculation
    // Shows how to set Aspose.Cells to manual calculation mode, change a source value, and run CalculateFormula with a custom CalculationMonitor that logs only the cells whose values actually changed. Includes an initial full calculation, dependent updates, and saving without automatic recalculation.

namespace AsposeCellsExamples
{
    // Custom monitor to report cells whose values changed after calculation
    public class ChangedCellMonitor : AbstractCalculationMonitor
    {
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Report only cells whose value actually changed

            // ValueChanged is true only when the cell's value differs from the previous value
            if (ValueChanged)
            {
                Console.WriteLine($"Cell {CellReference(rowIndex, colIndex)} changed from [{OriginalValue}] to [{CalculatedValue}]");
            }
        }

        // Helper to convert row/column indexes to A1 style reference
        private string CellReference(int row, int col)
        {

        private string CellReference(int row, int col)
        {
            // Convert zero‑based row/col to A1 style reference
            return CellsHelper.CellIndexToName(row, col);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and set up data/formulas
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Source values
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Formulas that depend on the source values
            cells["B1"].Formula = "=A1 + A2";          // Sum of A1 and A2
            cells["C1"].Formula = "=B1 * 2";           // Double the sum
            cells["D1"].Formula = "=C1 + 5";           // Add constant

            // -------------------------------------------------
            // 2. Perform the initial calculation (full)
            // -------------------------------------------------
            workbook.CalculateFormula();

            Console.WriteLine("Initial calculation results:");
            Console.WriteLine($"B1 = {cells["B1"].Value}");
            Console.WriteLine($"C1 = {cells["C1"].Value}");
            Console.WriteLine($"D1 = {cells["D1"].Value}");
            Console.WriteLine();

            // -------------------------------------------------
            // 3. Change a source cell (A1) – only dependent cells should recalc
            // -------------------------------------------------
            cells["A1"].PutValue(30); // Modify source value

            // Set calculation mode to Manual to avoid automatic recalculation on save
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Prepare calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new ChangedCellMonitor(),
                // Keep default Recursive = true so dependents are updated
                // IgnoreError = false (default) – keep errors visible
            };

            // -------------------------------------------------
            // 4. Recalculate – only cells whose values changed will be reported
            // -------------------------------------------------
            workbook.CalculateFormula(options);

            Console.WriteLine();
            Console.WriteLine("After modifying A1 and recalculating:");
            Console.WriteLine($"B1 = {cells["B1"].Value}");
            Console.WriteLine($"C1 = {cells["C1"].Value}");
            Console.WriteLine($"D1 = {cells["D1"].Value}");

            // -------------------------------------------------
            // 5. Save the workbook (no automatic recalculation on save)
            // -------------------------------------------------
            workbook.Save("RecalcDemo.xlsx");
        }
    }
}
