// Title: Interrupt Aspose.Cells formula calculation at cell B2 with a custom BeforeCalculate monitor in C#
// AI Prompts: Derive a class from AbstractCalculationMonitor, override BeforeCalculate, and call Interrupt() when the sheet, row, and column match the target cell. | Create an InterruptMonitor, assign it to a Workbook, set CalculationOptions.CalculationMonitor to your custom monitor, and run Workbook.CalculateFormula to stop processing at the specified cell. | Extend the monitor to examine the cell's value and trigger Interrupt() only when a condition (e.g., value > 100) is satisfied.
// Common Searches: Aspose.Cells C# interrupt calculation when cell B2 is evaluated | How to use AbstractCalculationMonitor to stop formula evaluation in Aspose.Cells | C# example of InterruptMonitor with CalculationOptions in Aspose.Cells | Conditional formula interruption based on cell address Aspose.Cells .NET | Stop workbook calculation at a specific cell using Aspose.Cells API
// Tags: Aspose.Cells custom calculation monitor | InterruptMonitor usage C# | BeforeCalculate cell address interruption | Conditional formula calculation stop Aspose.Cells | Workbook.CalculateFormula interrupt

using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Custom calculation monitor that can interrupt the calculation
    // The example shows how to halt Aspose.Cells formula calculation at a designated cell (B2) by implementing a ConditionalInterruptMonitor that inherits AbstractCalculationMonitor. The monitor overrides BeforeCalculate, checks the current sheet, row, and column, and calls Interrupt() via an InterruptMonitor. The monitor is attached to the workbook through CalculationOptions, causing Workbook.CalculateFormula to throw an Interrupted exception when the target cell is reached, after which the workbook can be saved or further processed.
    public class ConditionalInterruptMonitor : AbstractCalculationMonitor
    {
        private readonly InterruptMonitor _interruptMonitor;
        private readonly int _targetSheetIndex;
        private readonly int _targetRowIndex;
        private readonly int _targetColumnIndex;

        // Constructor receives the interrupt monitor and the cell to watch
        public ConditionalInterruptMonitor(InterruptMonitor interruptMonitor,
                                           int sheetIndex, int rowIndex, int columnIndex)
        {
            _interruptMonitor = interruptMonitor;
            _targetSheetIndex = sheetIndex;
            _targetRowIndex = rowIndex;
            _targetColumnIndex = columnIndex;
        }

        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // If the current cell matches the predefined address, request interruption
            if (sheetIndex == _targetSheetIndex &&
                rowIndex == _targetRowIndex &&
                colIndex == _targetColumnIndex)
            {
                // Interrupt the ongoing calculation
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
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some data and formulas
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].Formula = "=A1+A2";   // Simple sum
                sheet.Cells["B2"].Formula = "=A3*2";    // Cell we will monitor (B2)

                // 2. Set up an interrupt monitor and assign it to the workbook
                InterruptMonitor interruptMonitor = new InterruptMonitor();
                workbook.InterruptMonitor = interruptMonitor;

                // 3. Create a calculation monitor that interrupts when B2 is about to be calculated
                // B2 corresponds to row index 1, column index 1 (zero‑based)
                ConditionalInterruptMonitor calcMonitor = new ConditionalInterruptMonitor(
                    interruptMonitor,
                    sheetIndex: 0,
                    rowIndex: 1,
                    columnIndex: 1);

                // 4. Configure calculation options to use our monitor
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CalculationMonitor = calcMonitor
                };

                // 5. Perform calculation (the monitor will trigger interruption on B2)
                try
                {
                    workbook.CalculateFormula(calcOptions);
                    Console.WriteLine("Calculation completed without interruption.");
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
                {
                    Console.WriteLine("Calculation was interrupted as expected.");
                    // No need to change calculation mode; the workbook will remain in its current state.
                }

                // 6. Save the workbook (lifecycle rule: save)
                try
                {
                    workbook.Save("InterruptDemo.xlsx");
                    Console.WriteLine("Workbook saved successfully.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
