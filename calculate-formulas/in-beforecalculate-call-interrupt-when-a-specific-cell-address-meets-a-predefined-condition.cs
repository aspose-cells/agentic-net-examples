// Title: Interrupt Aspose.Cells Calculation for a Specific Cell Using ConditionalInterruptCalculationMonitor (C#)
// Description: This example shows how to create a custom calculation monitor that overrides BeforeCalculate, calls Interrupt() when a predefined cell address is reached, and attaches the monitor via CalculationOptions to Workbook.CalculateFormula. The resulting CellsException.Interrupted is caught, allowing you to stop formula evaluation for a target cell (e.g., B2) and then save the workbook.
// Keywords: Aspose.Cells | InterruptMonitor | ConditionalInterruptCalculationMonitor | BeforeCalculate | CalculateFormula interruption | CellsException.Interrupted | C# workbook calculation stop | custom calculation monitor
// Common Searches: Aspose.Cells stop calculation for a cell | How to use InterruptMonitor in Aspose.Cells .NET | ConditionalInterruptCalculationMonitor example | Catch CellsException.Interrupted during CalculateFormula | Abort formula evaluation in Aspose.Cells
// Developer Intent: Stop the formula calculation process when a specific cell is about to be evaluated.
// Use Cases: Terminate calculation when a high‑cost formula is reached. | Prevent runtime errors by halting evaluation of a cell that may cause divide‑by‑zero or overflow. | Implement validation that aborts processing if a key cell fails a business rule. | Improve performance in large workbooks by skipping unwanted calculations.
// AI Prompts: Generate code to log the sheet, row, and column inside BeforeCalculate before interrupting. | Show how to resume workbook calculation after handling the interruption exception. | Adapt the monitor to interrupt based on a cell's value threshold instead of its address. | Explain how to attach multiple calculation monitors for different cells. | Provide a unit test for ConditionalInterruptCalculationMonitor.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Custom calculation monitor that interrupts when a specific cell is about to be calculated
    // This example shows how to create a custom calculation monitor that overrides BeforeCalculate, calls Interrupt() when a predefined cell address is reached, and attaches the monitor via CalculationOptions to Workbook.CalculateFormula. The resulting CellsException.Interrupted is caught, allowing you to stop formula evaluation for a target cell (e.g., B2) and then save the workbook.
    public class ConditionalInterruptCalculationMonitor : AbstractCalculationMonitor
    {
        private readonly InterruptMonitor _interruptMonitor;
        private readonly int _targetSheetIndex;
        private readonly int _targetRowIndex;
        private readonly int _targetColumnIndex;

        public ConditionalInterruptCalculationMonitor(
            InterruptMonitor interruptMonitor,
            int targetSheetIndex,
            int targetRowIndex,
            int targetColumnIndex)
        {
            _interruptMonitor = interruptMonitor;
            _targetSheetIndex = targetSheetIndex;
            _targetRowIndex = targetRowIndex;
            _targetColumnIndex = targetColumnIndex;
        }

        // Called before each cell calculation
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Interrupt when the target cell is about to be calculated
            if (sheetIndex == _targetSheetIndex &&
                rowIndex == _targetRowIndex &&
                colIndex == _targetColumnIndex)
            {
                _interruptMonitor.Interrupt();
            }
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

                // Populate some data
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(10);

                // Set a formula in B2 (the cell we want to monitor)
                sheet.Cells["B2"].Formula = "=A1+A2";

                // Create an interrupt monitor and assign it to the workbook
                InterruptMonitor interruptMonitor = new InterruptMonitor();
                workbook.InterruptMonitor = interruptMonitor;

                // Create the custom calculation monitor targeting cell B2 (row 1, column 1, zero‑based)
                var calcMonitor = new ConditionalInterruptCalculationMonitor(
                    interruptMonitor,
                    targetSheetIndex: 0,
                    targetRowIndex: 1,    // B2 row index (zero‑based)
                    targetColumnIndex: 1); // B2 column index (zero‑based)

                // Set calculation options with the custom monitor
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CalculationMonitor = calcMonitor
                };

                // Perform calculation; expect an interruption when B2 is processed
                try
                {
                    workbook.CalculateFormula(calcOptions);
                    Console.WriteLine("Calculation completed without interruption (unexpected).");
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
                {
                    Console.WriteLine("Calculation was interrupted as intended.");
                }

                // Ensure the output directory exists
                string outputPath = "Result.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
