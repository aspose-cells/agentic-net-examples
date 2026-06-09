using System;
using Aspose.Cells;

namespace AsposeCellsFormulaTriggerDemo
{
    // Custom monitor that reacts when a specific cell's formula (or value) changes.
    // In this example we watch cell B1 (row 0, column 1).
    public class SpecificCellChangeMonitor : AbstractFormulaChangeMonitor
    {
        private readonly Workbook _workbook;
        private readonly int _targetRow;
        private readonly int _targetColumn;

        public SpecificCellChangeMonitor(Workbook workbook, int targetRow, int targetColumn)
        {
            _workbook = workbook;
            _targetRow = targetRow;
            _targetColumn = targetColumn;
        }

        // This method is called by the user (or by Aspose.Cells operations) when a cell's
        // formula/value changes. If the changed cell matches the target, we recalculate.
        public override void OnCellFormulaChanged(int sheetIndex, int rowIndex, int columnIndex)
        {
            if (rowIndex == _targetRow && columnIndex == _targetColumn)
            {
                Console.WriteLine($"Target cell changed at Sheet{sheetIndex}!{CellsHelper.CellIndexToName(rowIndex, columnIndex)}. Recalculating workbook...");
                _workbook.CalculateFormula();
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook.
            Workbook workbook = new Workbook();

            // 2. Set calculation mode to Manual so formulas are not auto‑recalculated.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // 3. Prepare a simple dependency: A1 = B1 * 2
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "=B1*2";

            // 4. Set an initial value for B1.
            sheet.Cells["B1"].PutValue(5);

            // 5. Create the monitor that watches B1 (row 0, column 1).
            var monitor = new SpecificCellChangeMonitor(workbook, 0, 1);

            // 6. Simulate a change to B1.
            sheet.Cells["B1"].PutValue(10);

            // 7. Manually notify the monitor about the change.
            //    In a real scenario this could be triggered by InsertRows/DeleteRows etc.
            monitor.OnCellFormulaChanged(0, 0, 1); // sheetIndex=0, row=0, column=1 (B1)

            // 8. Verify that A1 has been updated after the recalculation.
            Console.WriteLine($"A1 value after recalculation: {sheet.Cells["A1"].Value}");

            // 9. Save the workbook (using the standard lifecycle rule).
            workbook.Save("FormulaTriggerResult.xlsx");
        }
    }
}