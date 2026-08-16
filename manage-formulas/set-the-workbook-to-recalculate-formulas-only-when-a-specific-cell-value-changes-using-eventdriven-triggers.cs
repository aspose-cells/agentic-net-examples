// Title: Trigger formula recalculation only on a specific cell change with Aspose.Cells in C#
// Description: Demonstrates how to set Aspose.Cells to Manual calculation mode and use a custom class derived from AbstractFormulaChangeMonitor to recalculate formulas only when a designated cell (e.g., B1) is modified, improving performance for large workbooks.
// Keywords: Aspose.Cells manual calculation mode | C# formula change monitor | AbstractFormulaChangeMonitor example | specific cell trigger recalculation | event‑driven formula update | Excel workbook performance | CalculateFormula on cell change | Aspose.Cells custom monitor | recalculate only when B1 changes
// Common Searches: Aspose.Cells recalculate formulas only when a cell changes | C# manual calculation mode Aspose.Cells | How to use AbstractFormulaChangeMonitor | Trigger CalculateFormula after updating B1 | Event‑driven formula recalculation in .NET
// Developer Intent: Recalculate workbook formulas only when a designated cell changes.
// Use Cases: Reduce unnecessary calculations in large spreadsheets by switching to Manual mode and invoking CalculateFormula only after key input cells are edited. | Create reusable monitors for any target cell by adjusting row and column constants in the SpecificCellMonitor class. | Integrate the monitor into data‑entry applications where formulas must refresh only after user‑provided values are confirmed.
// AI Prompts: Generate a C# class that inherits from AbstractFormulaChangeMonitor to recalculate formulas when cell C3 changes. | Show how to enable Manual calculation mode in Aspose.Cells and programmatically trigger CalculateFormula after updating a range of cells. | Explain how to register multiple SpecificCellMonitor instances to watch several trigger cells in a workbook.

using System;
using Aspose.Cells;

// Demonstrates how to set Aspose.Cells to Manual calculation mode and use a custom class derived from AbstractFormulaChangeMonitor to recalculate formulas only when a designated cell (e.g., B1) is modified, improving performance for large workbooks.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set calculation mode to Manual so formulas are not auto‑recalculated
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Example formula that depends on cell B1
        sheet.Cells["A1"].Formula = "=B1*2";

        // Initial value for the trigger cell (B1)
        sheet.Cells["B1"].PutValue(5);

        // Create a monitor that will recalculate only when B1 changes
        var monitor = new SpecificCellMonitor(workbook);

        // Change the value of B1
        sheet.Cells["B1"].PutValue(10);

        // Notify the monitor that B1 (row 0, column 1) has changed
        monitor.OnCellFormulaChanged(0, 0, 1); // sheetIndex, rowIndex, columnIndex

        // Save the workbook
        workbook.Save("output.xlsx");
    }

    // Custom monitor that reacts only to a specific cell change
    class SpecificCellMonitor : AbstractFormulaChangeMonitor
    {
        private readonly Workbook _workbook;
        private const int TargetRow = 0;      // Row index for B1
        private const int TargetColumn = 1;   // Column index for B1

        public SpecificCellMonitor(Workbook workbook)
        {
            _workbook = workbook;
        }

        public override void OnCellFormulaChanged(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Recalculate only when the monitored cell (B1) changes
            if (rowIndex == TargetRow && columnIndex == TargetColumn)
            {
                _workbook.CalculateFormula();
                Console.WriteLine($"Recalculated because cell {CellsHelper.CellIndexToName(rowIndex, columnIndex)} changed.");
            }
        }
    }
}
