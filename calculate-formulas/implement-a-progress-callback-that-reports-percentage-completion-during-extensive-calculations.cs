// Title: Implement a C# Aspose.Cells calculation monitor that logs percentage progress of formula evaluation
// AI Prompts: Create a C# class that inherits from AbstractCalculationMonitor and outputs the current cell index and percentage completed before each formula is calculated. | Write code to enumerate all formula cells in a workbook, initialize CalculationOptions with the custom monitor, and invoke workbook.CalculateFormula to display real‑time progress. | Modify the ProgressCalculationMonitor to raise a .NET event instead of writing to the console, enabling UI components to receive progress updates.
// Common Searches: how to monitor formula calculation progress in Aspose.Cells using C# | Aspose.Cells calculation monitor example showing percentage completed | C# progress callback for workbook.CalculateFormula in Aspose.Cells | count total formula cells before running CalculateFormula Aspose.Cells | handle circular references in custom calculation monitor Aspose.Cells
// Tags: custom AbstractCalculationMonitor implementation | formula calculation progress reporting Aspose.Cells | percentage completion callback workbook.CalculateFormula | total formula cell count Aspose.Cells | circular reference handling calculation monitor

using System;
using Aspose.Cells;
using System.Collections;

namespace AsposeCellsProgressDemo
{
    // Custom monitor that reports calculation progress as a percentage.
    // The example defines a ProgressCalculationMonitor class derived from AbstractCalculationMonitor that increments a processed‑cell counter in BeforeCalculate, computes the completion percentage based on the total number of formula cells, and writes the progress to the console. The program counts all formula cells in the workbook, assigns the monitor to CalculationOptions, runs workbook.CalculateFormula with live progress reporting, and finally saves the workbook.
    public class ProgressCalculationMonitor : AbstractCalculationMonitor
    {
        private readonly int _totalFormulaCells;
        private int _processedCells;

        public ProgressCalculationMonitor(int totalFormulaCells)
        {
            _totalFormulaCells = totalFormulaCells > 0 ? totalFormulaCells : 1; // avoid division by zero
            _processedCells = 0;
        }

        // Called before each cell is calculated.
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            _processedCells++;
            int percent = (int)((double)_processedCells / _totalFormulaCells * 100);
            Console.WriteLine($"Calculating cell {_processedCells}/{_totalFormulaCells} ({percent}%) - Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
        }

        // Optional: after calculation you could also output details.
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // No additional action needed for this demo.
        }

        // Optional: handle circular references if they occur.
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected during calculation.");
            return true; // continue calculation
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and populate it with many formulas to simulate an extensive calculation.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill 1000 rows with simple dependent formulas.
            int rowCount = 1000;
            cells["A1"].PutValue(1);
            for (int i = 2; i <= rowCount; i++)
            {
                // Each cell adds the previous cell value plus 1.
                cells[$"A{i}"].Formula = $"=A{i - 1}+1";
            }

            // Add a few additional formulas across other columns.
            for (int i = 1; i <= rowCount; i++)
            {
                cells[$"B{i}"].Formula = $"=A{i}*2";
                cells[$"C{i}"].Formula = $"=SUM(A{i}:B{i})";
            }

            // Count total formula cells to calculate percentage.
            int totalFormulaCells = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Cell cell in ws.Cells)
                {
                    if (cell.IsFormula)
                        totalFormulaCells++;
                }
            }

            // Set up calculation options with the custom progress monitor.
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new ProgressCalculationMonitor(totalFormulaCells)
            };

            // Perform the calculation while the monitor reports progress.
            workbook.CalculateFormula(options);

            // Save the workbook (the save operation is not part of the progress monitoring).
            workbook.Save("ProgressCalculationDemo.xlsx");
        }
    }
}
