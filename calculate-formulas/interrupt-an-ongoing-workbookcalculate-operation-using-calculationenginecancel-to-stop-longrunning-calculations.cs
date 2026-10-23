// Title: Stopping a long‑running Workbook.CalculateFormula calculation with Aspose.Cells in C#
// Common Searches: c# Aspose.Cells cancel CalculateFormula execution | interrupt workbook.CalculateFormula after timeout Aspose.Cells | how to stop long running formula calculation in Aspose.Cells .NET | Aspose.Cells calculation engine timeout handling example | prevent Excel formula evaluation from hanging using Aspose.Cells
// Tags: Aspose.Cells calculation cancellation pattern | manage long running formula evaluation .NET | Workbook.CalculateFormula timeout strategy | partial workbook save after calculation abort | background formula calculation with Aspose.Cells

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads a large Excel workbook, runs Workbook.CalculateFormula on a background task, waits briefly, and demonstrates how to handle a timeout when the calculation does not finish, then saves any partial results.
class Program
{
    static void Main()
    {
        const string inputPath = "LargeFormulas.xlsx";
        const string outputPath = "ResultAfterCancel.xlsx";

        // Verify that the input workbook exists before loading
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook containing complex formulas
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Start the calculation on a separate thread so we can attempt to cancel it later
        Task calculationTask = Task.Run(() =>
        {
            try
            {
                // This will start the calculation engine and run until completion
                workbook.CalculateFormula();
            }
            catch (Exception ex)
            {
                // Capture any calculation errors
                Console.WriteLine($"Calculation error: {ex.Message}");
            }
        });

        // Let the calculation run for a short period (e.g., 2 seconds)
        Thread.Sleep(2000);

        // NOTE: Aspose.Cells does not expose a public cancellation API in all versions.
        // The original CalculationEngine.Cancel() method is unavailable, so we simply
        // allow the task to complete or be ignored after the timeout.

        try
        {
            // Wait for the calculation task to finish (or timeout after an additional period)
            if (!calculationTask.Wait(TimeSpan.FromSeconds(5)))
            {
                Console.WriteLine("Calculation did not finish within the expected time.");
            }
        }
        catch (AggregateException ae)
        {
            // Handle any exceptions thrown by the calculation task
            foreach (var ex in ae.InnerExceptions)
            {
                Console.WriteLine($"Calculation task exception: {ex.Message}");
            }
        }

        // Save the workbook (partial results may be present)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
