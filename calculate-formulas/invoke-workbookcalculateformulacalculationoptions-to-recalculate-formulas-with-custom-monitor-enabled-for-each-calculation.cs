// Title: C# – Recalculate Workbook Formulas with a Custom Calculation Monitor in Aspose.Cells
// Description: Demonstrates how to attach a user‑defined SampleCalculationMonitor (derived from AbstractCalculationMonitor) to CalculationOptions, then invoke Workbook.CalculateFormula to recalculate all formulas while logging cell coordinates, original and calculated values, and handling circular references. The workbook is saved after the calculation.
// Keywords: Aspose.Cells | C# | .NET | Workbook.CalculateFormula | CalculationOptions | AbstractCalculationMonitor | custom calculation monitor | formula recalculation logging | circular reference detection | track cell calculation progress
// Common Searches: Aspose.Cells custom calculation monitor example | How to use CalculationOptions with Workbook.CalculateFormula | C# monitor formula calculation Aspose.Cells | Detect circular references during workbook recalculation | Log before and after cell calculation Aspose.Cells
// Developer Intent: Implement a custom monitor to observe and control formula evaluation when recalculating a workbook with Workbook.CalculateFormula.
// Use Cases: Log each cell’s sheet, row, and column before and after evaluation for debugging or auditing. | Capture original and newly calculated values to verify formula correctness. | Detect circular references during recalculation and decide whether to continue processing.
// AI Prompts: Write C# code that creates an AbstractCalculationMonitor subclass to log cell positions and handle circular references, then use it with Workbook.CalculateFormula. | Explain step‑by‑step how to configure CalculationOptions with a custom monitor for formula recalculation in Aspose.Cells. | Show how to retrieve OriginalValue, CalculatedValue, and ValueChanged inside the AfterCalculate method of a custom calculation monitor.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor to track calculation progress
    // Demonstrates how to attach a user‑defined SampleCalculationMonitor (derived from AbstractCalculationMonitor) to CalculationOptions, then invoke Workbook.CalculateFormula to recalculate all formulas while logging cell coordinates, original and calculated values, and handling circular references. The workbook is saved after the calculation.
    public class SampleCalculationMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
            Console.WriteLine($"Original Value: {OriginalValue}, Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
        }

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            // Continue calculation for circular cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set sample formulas
            sheet.Cells["A1"].Formula = "=1+2";
            sheet.Cells["A2"].Formula = "=A1*3";
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Configure calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new SampleCalculationMonitor(),
                // Optional: other options can be set here
                IgnoreError = false,
                Recursive = true
            };

            // Recalculate all formulas using the options
            workbook.CalculateFormula(options);

            // Output the calculated values
            Console.WriteLine($"A1 = {sheet.Cells["A1"].Value}");
            Console.WriteLine($"A2 = {sheet.Cells["A2"].Value}");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");

            // Save the workbook (using the standard save rule)
            workbook.Save("CalculationMonitorResult.xlsx");
        }
    }
}
