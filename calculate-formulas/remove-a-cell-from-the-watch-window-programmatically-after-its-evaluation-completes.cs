using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put some data and a formula that will be watched
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].Formula = "=A1*2";

        // Add the cell to the watch window (watching A2)
        int watchIndex = sheet.CellWatches.Add("A2");

        // Set up calculation options with a custom monitor
        CalculationOptions options = new CalculationOptions();
        options.CalculationMonitor = new WatchRemovalMonitor(sheet);

        // Calculate formulas; AfterCalculate will be invoked for each calculated cell
        workbook.CalculateFormula(options);

        // Save the workbook (the watch for A2 will have been removed)
        workbook.Save("WatchRemoved.xlsx");
    }

    // Custom calculation monitor that removes a cell from the watch window after it is evaluated
    class WatchRemovalMonitor : AbstractCalculationMonitor
    {
        private readonly Worksheet _worksheet;

        public WatchRemovalMonitor(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Ensure the callback is for the worksheet we are monitoring
            if (sheetIndex != _worksheet.Index)
                return;

            // Locate the watch that corresponds to the calculated cell
            CellWatchCollection watches = _worksheet.CellWatches;
            for (int i = 0; i < watches.Count; i++)
            {
                CellWatch watch = watches[i];
                if (watch.Row == rowIndex && watch.Column == colIndex)
                {
                    // Remove the watch using the collection's RemoveAt method
                    watches.RemoveAt(i);
                    break; // Exit after removal
                }
            }
        }
    }
}