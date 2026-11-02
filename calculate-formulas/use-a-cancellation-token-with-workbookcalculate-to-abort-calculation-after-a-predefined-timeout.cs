// Title: Cancel Aspose.Cells workbook.CalculateFormula with a CancellationToken after a timeout (C#)
// Common Searches: c# set timeout for Aspose.Cells workbook.CalculateFormula | using CancellationToken to stop Aspose.Cells formula calculation | example of aborting long‑running Excel calculation in Aspose.Cells .NET | how to handle timeout when calling workbook.CalculateFormula in Aspose.Cells | cancel workbook.CalculateFormula task with CancellationTokenSource
// Tags: cancellation token workbook.CalculateFormula | Aspose.Cells formula calculation timeout | abort long-running Excel calculation .NET | task‑based calculation cancellation Aspose.Cells | timeout handling Aspose.Cells workbook

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, creates a CancellationTokenSource that triggers after 5 seconds, runs workbook.CalculateFormula inside a Task, waits for the task with the same timeout, and handles cancellation, errors, or successful completion before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
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

        // Create a cancellation token source that cancels after 5 seconds
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

        try
        {
            // Run calculation in a separate task to allow timeout handling
            Task calcTask = Task.Run(() => workbook.CalculateFormula(), cts.Token);

            // Wait for completion or timeout
            if (!calcTask.Wait(TimeSpan.FromSeconds(5)))
            {
                Console.WriteLine("Workbook calculation timed out.");
                // Optionally, you could attempt to cancel the task if supported
                // cts.Cancel();
            }
            else if (calcTask.IsFaulted && calcTask.Exception != null)
            {
                // Rethrow the inner exception for clearer error reporting
                throw calcTask.Exception.InnerException ?? calcTask.Exception;
            }
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Workbook calculation was cancelled.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during calculation: {ex.Message}");
        }

        try
        {
            // Save the workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
