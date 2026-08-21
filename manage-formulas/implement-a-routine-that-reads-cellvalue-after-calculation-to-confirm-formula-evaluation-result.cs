// Title: Read Cell.Value After Formula Calculation with CustomCalculationMonitor in Aspose.Cells for .NET
// Description: Shows how to attach a CustomCalculationMonitor to a workbook, run Workbook.CalculateFormula, and read the Cell.Value of a formula cell to verify the evaluated result. The C# example logs original and new values and demonstrates re‑calculation after source data changes.
// Keywords: Aspose.Cells | C# | .NET | CustomCalculationMonitor | AbstractCalculationMonitor | Workbook.CalculateFormula | Cell.Value | formula evaluation | read calculated value | monitor cell changes
// Common Searches: Aspose.Cells read calculated cell value | CustomCalculationMonitor example C# | How to get formula result after CalculateFormula | Verify formula evaluation Aspose.Cells .NET | Log cell value changes during calculation
// Developer Intent: Retrieve the Cell.Value after invoking Workbook.CalculateFormula to confirm that a formula has been evaluated correctly.
// Use Cases: Audit spreadsheet modifications by logging original and new values for each calculated cell. | Automated unit tests that compare Cell.Value with expected outcomes after formula evaluation. | Trigger downstream business logic when source cells change, using the monitor to capture updated values.
// AI Prompts: Generate C# code that uses CustomCalculationMonitor to log OriginalValue and CalculatedValue for each cell during Aspose.Cells calculation. | Provide a snippet that reads Cell.Value after Workbook.CalculateFormula and asserts it equals a specified number. | Explain how to reuse a calculation monitor to track formula updates when source data is modified in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaMonitorDemo
{
    // Custom monitor to capture calculation details for each cell
    // Shows how to attach a CustomCalculationMonitor to a workbook, run Workbook.CalculateFormula, and read the Cell.Value of a formula cell to verify the evaluated result. The C# example logs original and new values and demonstrates re‑calculation after source data changes.
    public class CustomCalculationMonitor : AbstractCalculationMonitor
    {
        // This method is called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Check if the cell's value actually changed during calculation
            if (ValueChanged)
            {
                // OriginalValue – value before calculation
                // CalculatedValue – newly calculated value (the one we want to confirm)
                Console.WriteLine($"Cell ({rowIndex}, {columnIndex}) changed from [{OriginalValue}] to [{CalculatedValue}]");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Set up sample data and formulas
            //    A1 = 10 (plain value)
            //    A2 = =A1*3 (formula that depends on A1)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].Formula = "=A1*3";

            // 3. Attach the custom calculation monitor
            CustomCalculationMonitor monitor = new CustomCalculationMonitor();
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = monitor
            };

            // 4. Trigger calculation – the monitor will output change information
            workbook.CalculateFormula(options);

            // 5. After calculation, read the cell's Value property to confirm the result
            Cell resultCell = sheet.Cells["A2"];
            Console.WriteLine($"After calculation, cell A2 Value = {resultCell.Value}");

            // 6. (Optional) Change the source value and recalculate to see the monitor fire again
            sheet.Cells["A1"].PutValue(20);
            workbook.CalculateFormula(options);
            Console.WriteLine($"After second calculation, cell A2 Value = {resultCell.Value}");
        }
    }
}
