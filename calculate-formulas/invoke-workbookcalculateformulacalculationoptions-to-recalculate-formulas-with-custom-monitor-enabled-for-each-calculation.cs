// Title: Use a custom AbstractCalculationMonitor to log each cell during Workbook.CalculateFormula in Aspose.Cells for .NET
// AI Prompts: Write C# code that derives a SampleCalculationMonitor from AbstractCalculationMonitor, assigns it to CalculationOptions, and calls Workbook.CalculateFormula to output before/after messages for every cell. | Show how to implement the OnCircular method in a custom calculation monitor to detect circular references and continue processing while recalculating formulas with Aspose.Cells. | Create a complete example that sets sample formulas, attaches the monitor, recalculates the workbook, prints the resulting values, and saves the file as an .xlsx.
// Common Searches: attach a calculation monitor to Workbook.CalculateFormula in C# Aspose.Cells | log each cell's formula evaluation using Aspose.Cells | detect and continue circular references while recalculating formulas with Aspose.Cells .NET | monitor formula recalculation progress in an Aspose.Cells workbook | save workbook after formula recalculation with a custom monitor Aspose.Cells
// Tags: Aspose.Cells calculation callbacks implementation | formula recalculation progress tracking Aspose.Cells | cell calculation event logging C# | circular reference handling Aspose.Cells | export workbook after formula recalculation Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor to track formula calculation progress
    // The example creates a Workbook, adds three formulas, defines a SampleCalculationMonitor that logs before and after each cell calculation and detects circular references, assigns this monitor via CalculationOptions, recalculates all formulas with Workbook.CalculateFormula, prints the computed values, and saves the workbook as CalculationMonitorResult.xlsx.
    public class SampleCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        // Called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
            Console.WriteLine($"Original Value: {OriginalValue}, Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
        }

        // Called when a circular reference is detected
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

            // Create calculation options and assign the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new SampleCalculationMonitor()
            };

            // Recalculate all formulas with monitoring enabled
            workbook.CalculateFormula(options);

            // Output the calculated values
            Console.WriteLine($"A1 = {sheet.Cells["A1"].Value}");
            Console.WriteLine($"A2 = {sheet.Cells["A2"].Value}");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");

            // Save the workbook (optional)
            workbook.Save("CalculationMonitorResult.xlsx");
        }
    }
}
