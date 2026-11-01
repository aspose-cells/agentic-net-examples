// Title: Use a CancellationToken to abort Aspose.Cells workbook.CalculateFormula after a 10‑second timeout in C#
// Common Searches: c# how to stop Aspose.Cells workbook.CalculateFormula after a timeout | using CancellationTokenSource with Aspose.Cells CalculateFormula method | abort Excel formula calculation in Aspose.Cells when it takes too long | set timeout for workbook.CalculateFormula in Aspose.Cells C# example | handle OperationCanceledException while calculating formulas with Aspose.Cells
// Tags: Aspose.Cells calculation timeout with CancellationToken | Workbook.CalculateFormula cancellation in .NET | Excel formula evaluation abort using CancellationToken | C# task‑based timeout for Aspose.Cells calculation | OperationCanceledException handling in Aspose.Cells

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads an Excel workbook, creates a CancellationTokenSource with a 10‑second timeout, runs workbook.CalculateFormula on a background task, waits for completion or cancellation, catches OperationCanceledException and other errors, reports the result, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Set a timeout of 10 seconds for the calculation
        using (CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(10)))
        {
            bool completed = false;
            try
            {
                // Run calculation on a separate task to enable timeout handling
                Task calcTask = Task.Run(() => workbook.CalculateFormula(), cts.Token);

                // Wait for the task to complete or be cancelled
                calcTask.Wait(cts.Token);
                completed = true;
            }
            catch (OperationCanceledException)
            {
                // Calculation was aborted due to timeout
                Console.WriteLine("Calculation was cancelled after the timeout period.");
            }
            catch (AggregateException ae)
            {
                // Unwrap any exceptions thrown during calculation
                foreach (var inner in ae.InnerExceptions)
                {
                    Console.WriteLine($"Calculation error: {inner.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error during calculation: {ex.Message}");
            }

            if (completed)
            {
                Console.WriteLine("Calculation completed successfully.");
            }
        }

        // Save the workbook after calculation
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
