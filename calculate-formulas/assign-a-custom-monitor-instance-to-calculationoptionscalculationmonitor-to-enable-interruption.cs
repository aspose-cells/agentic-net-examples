// Title: C# – Assign a Custom CalculationMonitor to CalculationOptions for Real‑Time Formula Tracking in Aspose.Cells
// Description: Learn how to monitor and interrupt workbook formula calculations in Aspose.Cells for .NET by creating a class that inherits AbstractCalculationMonitor, overriding BeforeCalculate, AfterCalculate, and OnCircular, assigning it to CalculationOptions.CalculationMonitor, and calling Workbook.CalculateFormula with the custom options.
// Keywords: Aspose.Cells | CalculationMonitor | CalculationOptions | AbstractCalculationMonitor | C# | .NET | formula calculation progress | interrupt calculation | circular reference handling | real‑time workbook monitoring | CalculateFormula
// Common Searches: Aspose.Cells custom CalculationMonitor example | How to use AbstractCalculationMonitor in C# | Assign CalculationMonitor to CalculationOptions Aspose | Track formula calculation progress Aspose.Cells | Interrupt long running calculations Aspose.Cells .NET
// Developer Intent: Implement a custom monitor to track, log, and optionally abort formula calculations by setting CalculationOptions.CalculationMonitor.
// Use Cases: Log start and end of each cell calculation for performance analysis. | Detect and handle circular references during workbook evaluation. | Update UI elements such as progress bars or status messages in real time. | Provide a user‑controlled cancel option for large or complex workbooks.
// AI Prompts: Create a SampleCalculationMonitor that stops calculation after 500 cells. | Show how to bind a WinForms ProgressBar to the BeforeCalculate and AfterCalculate callbacks. | Explain how to pause or cancel calculation based on a cancellation token supplied by the user.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor to track formula calculation progress
    // Learn how to monitor and interrupt workbook formula calculations in Aspose.Cells for .NET by creating a class that inherits AbstractCalculationMonitor, overriding BeforeCalculate, AfterCalculate, and OnCircular, assigning it to CalculationOptions.CalculationMonitor, and calling Workbook.CalculateFormula with the custom options.
    public class SampleCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculation: Sheet{sheetIndex}, Row{rowIndex}, Column{columnIndex}");
        }

        // Called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculation: Sheet{sheetIndex}, Row{rowIndex}, Column{columnIndex}");
            Console.WriteLine($"Original Value: {OriginalValue}, Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
        }

        // Handle circular references if needed
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            return base.OnCircular(circularCellsData);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set sample formulas
            worksheet.Cells["A1"].Formula = "=1+2";
            worksheet.Cells["A2"].Formula = "=A1*3";
            worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Create calculation options and assign the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new SampleCalculationMonitor()
            };

            // Perform calculation with monitoring
            workbook.CalculateFormula(options);

            // Output the results
            Console.WriteLine($"A1: {worksheet.Cells["A1"].Value}");
            Console.WriteLine($"A2: {worksheet.Cells["A2"].Value}");
            Console.WriteLine($"A3: {worksheet.Cells["A3"].Value}");

            // Save the workbook (optional)
            workbook.Save("CalculationMonitorResult.xlsx");
        }
    }
}
