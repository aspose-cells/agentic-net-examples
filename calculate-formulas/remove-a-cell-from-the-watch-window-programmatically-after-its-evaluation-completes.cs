// Title: How to automatically remove a cell watch from Aspose.Cells after its formula is calculated using C#
// AI Prompts: Implement a subclass of AbstractCalculationMonitor that detects when a specific cell finishes calculation and calls Worksheet.CellWatches.RemoveAt to delete the watch. | Show how to add a watch to a target cell, assign the custom monitor to CalculationOptions, and invoke Workbook.CalculateFormula so the watch is removed automatically. | Write C# code that captures the watch index and target cell coordinates, then removes the watch inside the AfterCalculate callback.
// Common Searches: Aspose.Cells C# remove cell watch after formula evaluation | custom calculation monitor to delete cell watch in .NET | programmatically clear Worksheet.CellWatches after calculation | example of using AbstractCalculationMonitor for watch cleanup in Aspose.Cells
// Tags: Aspose.Cells calculation monitor implementation | Worksheet.CellWatches removal programmatically | C# automatic cell watch cleanup | Aspose.Cells formula evaluation callback | using AbstractCalculationMonitor in .NET

using System;
using Aspose.Cells;

namespace AsposeCellsWatchRemovalDemo
{
    // Custom calculation monitor to remove a cell watch after its evaluation
    // The example defines a WatchRemovalMonitor that inherits from AbstractCalculationMonitor. In its AfterCalculate method it checks whether the evaluated cell matches the watched cell (B2) and, if so, removes the watch via Worksheet.CellWatches.RemoveAt. The program creates a workbook, adds a watch to B2, configures CalculationOptions with the custom monitor, runs Workbook.CalculateFormula, and saves the result, demonstrating automatic watch removal after the cell's formula is evaluated.
    class WatchRemovalMonitor : AbstractCalculationMonitor
    {
        private readonly Worksheet _worksheet;
        private readonly int _watchIndex;
        private readonly int _targetRow;
        private readonly int _targetColumn;

        public WatchRemovalMonitor(Worksheet worksheet, int watchIndex, int targetRow, int targetColumn)
        {
            _worksheet = worksheet;
            _watchIndex = watchIndex;
            _targetRow = targetRow;
            _targetColumn = targetColumn;
        }

        // This method is called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Check if the calculated cell is the one we are watching
            if (sheetIndex == _worksheet.Index && rowIndex == _targetRow && colIndex == _targetColumn)
            {
                // Remove the watch from the watch window
                _worksheet.CellWatches.RemoveAt(_watchIndex);
                Console.WriteLine($"Cell watch for {_worksheet.Cells[_targetRow, _targetColumn].Name} removed after calculation.");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put some sample data and formulas
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B2"].Formula = "=A1+A2"; // Cell to watch

            // Add a watch for cell B2 and get its index in the collection
            int watchIndex = sheet.CellWatches.Add("B2");

            // Determine row and column indices for B2 (0‑based)
            int targetRow = sheet.Cells["B2"].Row;      // 1
            int targetColumn = sheet.Cells["B2"].Column; // 1

            // Set up calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions();
            options.CalculationMonitor = new WatchRemovalMonitor(sheet, watchIndex, targetRow, targetColumn);

            // Perform calculation; the monitor will remove the watch after B2 is evaluated
            workbook.CalculateFormula(options);

            // Save the workbook (output file name can be changed as needed)
            workbook.Save("WatchRemovalResult.xlsx");
        }
    }
}
