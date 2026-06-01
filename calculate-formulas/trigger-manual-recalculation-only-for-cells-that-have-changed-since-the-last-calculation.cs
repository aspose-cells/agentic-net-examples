using System;
using Aspose.Cells;

class ManualRecalcDemo
{
    static void Main()
    {
        // ---------- Create ----------
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Set calculation mode to Manual so formulas are not auto‑recalculated
        wb.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Populate source cells and dependent formulas
        ws.Cells["A1"].PutValue(10);               // source value
        ws.Cells["A2"].PutValue(20);               // source value
        ws.Cells["B1"].Formula = "=A1+A2";         // depends on A1 and A2
        ws.Cells["C1"].Formula = "=B1*2";          // depends on B1

        // First manual calculation (calculates all formulas)
        wb.CalculateFormula();

        // ---------- Save ----------
        wb.Save("ManualRecalcDemo.xlsx");

        // ---------- Load ----------
        // Load the workbook we just saved
        Workbook loadedWb = new Workbook("ManualRecalcDemo.xlsx");
        Worksheet loadedWs = loadedWb.Worksheets[0];

        // Change only one source cell (A1)
        loadedWs.Cells["A1"].PutValue(30);

        // Set up a calculation monitor to report which cells actually changed
        var monitor = new ChangedCellMonitor();
        var options = new CalculationOptions
        {
            CalculationMonitor = monitor,
            // Recursive = true (default) ensures dependent cells are updated
            Recursive = true
        };

        // Recalculate – only cells whose value depends on the changed A1 will be updated
        loadedWb.CalculateFormula(options);

        // ---------- Save ----------
        loadedWb.Save("ManualRecalcDemo_Updated.xlsx");
    }

    // Custom monitor that prints changed cells after each calculation step
    class ChangedCellMonitor : AbstractCalculationMonitor
    {
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            if (ValueChanged)
            {
                string cellName = GetCellName(rowIndex, columnIndex);
                Console.WriteLine($"Cell {cellName} changed from [{OriginalValue}] to [{CalculatedValue}]");
            }
        }

        // Helper to convert zero‑based row/column indexes to Excel cell name (e.g., A1)
        private string GetCellName(int row, int column)
        {
            string colLetter = "";
            int dividend = column + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colLetter = Convert.ToChar('A' + modulo) + colLetter;
                dividend = (dividend - modulo) / 26;
            }
            return $"{colLetter}{row + 1}";
        }
    }
}