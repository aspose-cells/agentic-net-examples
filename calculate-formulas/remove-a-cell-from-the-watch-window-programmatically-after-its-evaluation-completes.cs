using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data and a formula that will be evaluated
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].Formula = "=A1*2";

        // Add the cell to the Watch Window (watching A2)
        sheet.CellWatches.Add("A2");

        // Set up calculation options with a custom monitor
        CalculationOptions options = new CalculationOptions();
        options.CalculationMonitor = new WatchRemovalMonitor(sheet);

        // Perform calculation; after the watched cell is calculated,
        // the monitor will remove it from the watch window
        workbook.CalculateFormula(options);

        // Save the workbook (the watch window will no longer contain A2)
        workbook.Save("WatchRemoved.xlsx");
    }

    // Custom calculation monitor that removes a cell from the watch window
    // after its evaluation completes.
    class WatchRemovalMonitor : AbstractCalculationMonitor
    {
        private readonly Worksheet _worksheet;

        public WatchRemovalMonitor(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Ensure we are handling the correct worksheet
            if (sheetIndex != _worksheet.Index) return;

            // Find the watch entry that matches the calculated cell
            CellWatchCollection watches = _worksheet.CellWatches;
            for (int i = 0; i < watches.Count; i++)
            {
                CellWatch watch = watches[i];
                if (watch.Row == rowIndex && watch.Column == colIndex)
                {
                    // Remove the watch entry
                    watches.RemoveAt(i);
                    break;
                }
            }
        }
    }
}