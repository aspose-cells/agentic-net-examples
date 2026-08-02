// Title: Read Cell.Value after formula calculation with a custom CalculationMonitor in Aspose.Cells for .NET
// Description: Shows how to attach an AbstractCalculationMonitor via CalculationOptions, execute Workbook.CalculateFormula, and read worksheet.Cells["A1"].Value and Cells["B1"].Value to confirm the computed results.
// Keywords: Aspose.Cells | C# | .NET | CalculationMonitor | AbstractCalculationMonitor | AfterCalculate | workbook.CalculateFormula | read cell value | formula evaluation | programmatic verification
// Common Searches: Aspose.Cells get cell value after CalculateFormula | How to use AbstractCalculationMonitor in C# | Read calculated result of a cell with Aspose.Cells | Log cell changes during formula recalculation Aspose | C# example for custom CalculationMonitor
// Developer Intent: Retrieve and verify the values of cells after running CalculateFormula, optionally logging changes through a custom monitor.
// Use Cases: Automated unit tests that compare expected and actual formula results. | Audit trails that capture original and new values when formulas recalculate. | Triggering business rules when a cell's calculated value differs from its previous state.
// AI Prompts: Generate C# code that calculates formulas with Aspose.Cells, then reads Cell.Value to compare with expected numbers. | Create an AbstractCalculationMonitor implementation that logs OriginalValue and CalculatedValue for each changed cell. | Show how to configure CalculationOptions with a custom monitor to capture AfterCalculate events and output the new values.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaMonitorDemo
{
    // Shows how to attach an AbstractCalculationMonitor via CalculationOptions, execute Workbook.CalculateFormula, and read worksheet.Cells["A1"].Value and Cells["B1"].Value to confirm the computed results.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set up formulas
            worksheet.Cells["A1"].Formula = "=1+2";      // Expected result: 3
            worksheet.Cells["B1"].Formula = "=A1*3";    // Expected result: 9

            // Create a custom calculation monitor
            var monitor = new MyCalculationMonitor();

            // Attach the monitor via calculation options
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = monitor
            };

            // Perform calculation – this will invoke AfterCalculate for each cell
            workbook.CalculateFormula(options);

            // After calculation, read the cell values directly to confirm results
            Console.WriteLine($"A1 value after calculation: {worksheet.Cells["A1"].Value}");
            Console.WriteLine($"B1 value after calculation: {worksheet.Cells["B1"].Value}");
        }

        // Custom monitor that reports changes using CalculatedValue and OriginalValue
        private class MyCalculationMonitor : AbstractCalculationMonitor
        {
            public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
            {
                // Only act when the cell's value actually changed
                if (ValueChanged)
                {
                    // CalculatedValue provides the newly computed value
                    Console.WriteLine($"Cell ({rowIndex},{columnIndex}) changed from [{OriginalValue}] to [{CalculatedValue}]");
                }
            }
        }
    }
}
