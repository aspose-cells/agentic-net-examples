using System;
using Aspose.Cells;

namespace AsposeCellsFormulaMonitorDemo
{
    // Custom monitor to capture calculation details for each cell
    public class FormulaMonitor : AbstractCalculationMonitor
    {
        // Store the last calculated value for demonstration purposes
        public object LastCalculatedValue { get; private set; }

        // This method is called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Check if the cell's value actually changed during calculation
            if (ValueChanged)
            {
                // OriginalValue is the value before calculation,
                // CalculatedValue is the newly computed value.
                Console.WriteLine($"Cell ({rowIndex}, {columnIndex}) changed from [{OriginalValue}] to [{CalculatedValue}]");
                LastCalculatedValue = CalculatedValue;
            }
            else
            {
                // No change – still useful for tracing
                Console.WriteLine($"Cell ({rowIndex}, {columnIndex}) unchanged. Value: [{OriginalValue}]");
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

            // 2. Populate cells: A1 with a static value, A2 with a formula that depends on A1
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].Formula = "=A1*2";

            // 3. Set up the calculation monitor and options
            FormulaMonitor monitor = new FormulaMonitor();
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = monitor
            };

            // 4. Trigger calculation – the monitor will be invoked for each calculated cell
            workbook.CalculateFormula(options);

            // 5. After calculation, read the cell's Value property to confirm the result
            Cell resultCell = sheet.Cells["A2"];
            Console.WriteLine($"After calculation, cell A2 Value = {resultCell.Value}");

            // 6. Optionally, verify that the monitor captured the same value
            Console.WriteLine($"Monitor captured CalculatedValue = {monitor.LastCalculatedValue}");

            // 7. Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("FormulaMonitorResult.xlsx");
        }
    }
}