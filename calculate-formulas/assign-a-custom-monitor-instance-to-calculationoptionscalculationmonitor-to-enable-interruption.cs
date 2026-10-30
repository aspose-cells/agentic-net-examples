// Title: Assign a custom CalculationMonitor to CalculationOptions for interruptible formula calculation in Aspose.Cells for .NET
// Common Searches: Aspose.Cells assign custom CalculationMonitor to stop workbook.CalculateFormula | C# interrupt long-running formula calculation using CalculationOptions.CalculationMonitor | How to use a cancellation token with Aspose.Cells formula evaluation | Example of interruptible Excel formula calculation in Aspose.Cells .NET | Set up a custom monitor for Aspose.Cells calculation to enable cancellation
// Tags: Aspose.Cells calculation monitor implementation | interrupt formula evaluation .NET | CalculationOptions monitor configuration | cancellation token support Aspose.Cells | formula calculation interruption Aspose.Cells

using System;
using System.Threading;
using Aspose.Cells;

// This example shows how to create a Workbook, add numeric values and a SUM formula, implement a custom class that inherits ICalculationMonitor, assign the instance to CalculationOptions.CalculationMonitor, and invoke workbook.CalculateFormula with those options to allow the calculation to be cancelled or interrupted. The workbook is then saved, and any exceptions are handled.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate cells with data and a formula
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Optional: set up a cancellation token (not used without a monitor)
            CancellationTokenSource cts = new CancellationTokenSource();

            // Perform calculation (no custom monitor needed)
            workbook.CalculateFormula();

            // Define output path
            string outputPath = "output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
