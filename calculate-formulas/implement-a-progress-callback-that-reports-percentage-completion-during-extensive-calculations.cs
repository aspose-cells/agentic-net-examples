// Title: Track Formula Calculation Progress with a Custom AbstractCalculationMonitor in Aspose.Cells for .NET
// Description: This example shows how to create a ProgressCalculationMonitor that inherits AbstractCalculationMonitor, counts all formula cells, and reports the completion percentage in the console during workbook.CalculateFormula. The monitor is attached via CalculationOptions, handles circular references, and the workbook is saved after processing.
// Keywords: Aspose.Cells | C# | .NET | CalculationMonitor | progress callback | formula calculation | percentage reporting | custom monitor | large workbook | circular reference handling
// Common Searches: Aspose.Cells progress monitor C# | how to show calculation percentage Aspose.Cells | custom AbstractCalculationMonitor example | track formula evaluation progress .NET | display workbook calculation progress in console
// Developer Intent: Implement a callback that reports the percentage of completed formula calculations while processing a large workbook.
// Use Cases: Log real‑time calculation progress to the console or a UI element. | Update a progress bar in WinForms/WPF during workbook.CalculateFormula. | Maintain calculation flow even when circular references are present.
// AI Prompts: Generate C# code that uses Aspose.Cells AbstractCalculationMonitor to update a WinForms progress bar while calculating formulas. | Write a method that scans a workbook, counts all formula cells, and initializes a custom calculation monitor with that count. | Explain how to modify ProgressCalculationMonitor to write progress data to a log file instead of the console.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsProgressDemo
{
    // Custom monitor that reports calculation progress as a percentage
    // This example shows how to create a ProgressCalculationMonitor that inherits AbstractCalculationMonitor, counts all formula cells, and reports the completion percentage in the console during workbook.CalculateFormula. The monitor is attached via CalculationOptions, handles circular references, and the workbook is saved after processing.
    public class ProgressCalculationMonitor : AbstractCalculationMonitor
    {
        private readonly int _totalCells;   // total number of cells to be calculated
        private int _processedCells;        // cells processed so far

        public ProgressCalculationMonitor(int totalCells)
        {
            _totalCells = totalCells > 0 ? totalCells : 1; // avoid division by zero
            _processedCells = 0;
        }

        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            _processedCells++;
            int percent = (int)((double)_processedCells / _totalCells * 100);
            Console.WriteLine($"Calculating Sheet{sheetIndex} Cell[{rowIndex}, {columnIndex}] - {percent}% completed");
        }

        // Optional: after calculation you can also display info
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Additional details can be output here if needed
        }

        // Optional: handle circular references
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Continue calculation even if circular references are found
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate the worksheet with a large number of formulas to simulate an extensive calculation
                int rowCount = 500; // adjust as needed for testing
                for (int i = 0; i < rowCount; i++)
                {
                    // Simple progressive formula: each cell sums the previous cell with a constant
                    if (i == 0)
                    {
                        cells[i, 0].PutValue(1); // A1 = 1
                    }
                    else
                    {
                        cells[i, 0].Formula = $"=A{i}+1"; // A{i+1} = A{i}+1
                    }
                }

                // Determine total number of formula cells that will be processed
                int totalFormulaCells = 0;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    Cells wsCells = ws.Cells;
                    int maxRow = wsCells.MaxDataRow;
                    int maxCol = wsCells.MaxDataColumn;
                    for (int r = 0; r <= maxRow; r++)
                    {
                        for (int c = 0; c <= maxCol; c++)
                        {
                            if (wsCells[r, c].IsFormula) // corrected property
                            {
                                totalFormulaCells++;
                            }
                        }
                    }
                }

                // Set up calculation options with the custom progress monitor
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CalculationMonitor = new ProgressCalculationMonitor(totalFormulaCells)
                };

                // Perform formula calculation with progress reporting
                workbook.CalculateFormula(calcOptions);

                // Save the workbook
                workbook.Save("ProgressCalculationDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
