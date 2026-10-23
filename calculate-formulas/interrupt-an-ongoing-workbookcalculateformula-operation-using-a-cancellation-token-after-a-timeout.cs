// Title: Handle timeout for Workbook.CalculateFormula in Aspose.Cells .NET using Task and CancellationToken
// Common Searches: set a timeout for Aspose.Cells Workbook.CalculateFormula in C# | detect and abort long‑running formula calculation with Aspose.Cells | use CancellationToken to stop CalculateFormula after a time limit in .NET | run Workbook.CalculateFormula on a background task with timeout handling | Aspose.Cells formula evaluation timeout example in C#
// Tags: Aspose.Cells CalculateFormula timeout handling | C# task-based cancellation for Excel formula calculation | Workbook.CalculateFormula long‑running operation management | Excel workbook formula evaluation abort .NET | Aspose.Cells background calculation with timeout

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, executes Workbook.CalculateFormula on a background task, waits up to five seconds, reports whether the calculation finished or timed out (noting that Aspose.Cells does not provide native cancellation), and then saves the workbook.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Perform formula calculation with a timeout (5 seconds)
        try
        {
            var calcTask = Task.Run(() => workbook.CalculateFormula());

            // Wait for the calculation to finish or timeout
            if (calcTask.Wait(TimeSpan.FromSeconds(5)))
            {
                Console.WriteLine("Formula calculation completed successfully.");
            }
            else
            {
                Console.WriteLine("Formula calculation timed out.");
                // Note: Aspose.Cells does not support cancellation of an ongoing calculation.
            }
        }
        catch (AggregateException ae)
        {
            // Unwrap and display any exceptions thrown during calculation
            Console.WriteLine($"Error during formula calculation: {ae.Flatten().InnerException?.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error during calculation: {ex.Message}");
        }

        // Save the workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
