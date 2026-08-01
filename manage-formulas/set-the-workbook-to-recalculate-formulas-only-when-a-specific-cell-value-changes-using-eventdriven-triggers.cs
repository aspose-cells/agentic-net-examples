// Title: Trigger Workbook Recalculation Only When a Specific Cell Changes in Aspose.Cells (C#)
// Description: Shows how to set Aspose.Cells to manual calculation mode and implement a custom AbstractFormulaChangeMonitor that recalculates the workbook only when a designated cell (e.g., A1) is modified, improving performance for large spreadsheets.
// Keywords: Aspose.Cells | C# | manual calculation mode | AbstractFormulaChangeMonitor | event‑driven formula recalculation | specific cell trigger | workbook.CalculateFormula | formula monitoring | performance optimization | spreadsheet automation
// Common Searches: Aspose.Cells recalculate formulas only when a cell changes | How to use AbstractFormulaChangeMonitor in .NET | Set manual calculation mode and trigger calculation on cell update Aspose.Cells | Event‑driven formula calculation C# Aspose.Cells | Optimize large workbook performance Aspose.Cells manual mode
// Developer Intent: Recalculate the workbook’s formulas only after a particular cell’s value is changed.
// Use Cases: Large financial models where only a key input cell (e.g., interest rate) should trigger a full recalculation. | Interactive dashboards that recalc formulas only when the user edits a configuration cell, reducing latency. | Batch processing pipelines that defer formula evaluation until a control cell signals that data is ready.
// AI Prompts: Generate a C# example that recalculates an Aspose.Cells workbook when cell B2 changes using AbstractFormulaChangeMonitor. | Explain how to monitor multiple cells and trigger separate recalculation actions in Aspose.Cells. | Show how to switch back to automatic calculation mode after using a custom cell‑change monitor.

using System;
using Aspose.Cells;

// Shows how to set Aspose.Cells to manual calculation mode and implement a custom AbstractFormulaChangeMonitor that recalculates the workbook only when a designated cell (e.g., A1) is modified, improving performance for large spreadsheets.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set calculation mode to Manual so formulas are not auto‑recalculated
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Sample formulas that depend on cell A1
        sheet.Cells["B1"].Formula = "=A1*2";
        sheet.Cells["C1"].Formula = "=B1+5";

        // Instantiate a monitor that watches a specific cell (A1)
        var monitor = new SpecificCellMonitor(workbook, "A1");

        // Change the value of the watched cell
        sheet.Cells["A1"].PutValue(10);

        // Manually notify the monitor that the cell's formula (or value) changed
        // Row 0, Column 0 correspond to A1
        monitor.OnCellFormulaChanged(0, 0, 0);

        // Save the workbook (lifecycle save)
        workbook.Save("output.xlsx");
    }

    // Custom monitor derived from AbstractFormulaChangeMonitor
    class SpecificCellMonitor : AbstractFormulaChangeMonitor
    {
        private readonly Workbook _workbook;
        private readonly string _targetAddress;

        public SpecificCellMonitor(Workbook workbook, string targetAddress)
        {
            _workbook = workbook;
            _targetAddress = targetAddress;
        }

        // Triggered when a cell's formula changes; we use it as a generic change event
        public override void OnCellFormulaChanged(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Convert indices to A1 style address
            string changedAddress = CellsHelper.CellIndexToName(rowIndex, columnIndex);

            // If the changed cell is the one we monitor, recalculate the workbook
            if (changedAddress.Equals(_targetAddress, StringComparison.OrdinalIgnoreCase))
            {
                _workbook.CalculateFormula();
                Console.WriteLine($"Workbook recalculated because {_targetAddress} changed.");
            }
        }
    }
}
