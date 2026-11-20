// Title: How to abort Aspose.Cells Workbook.CalculateFormula with a CancellationToken after a timeout in C#
// Common Searches: c# cancel Aspose.Cells CalculateFormula after 5 seconds | use CancellationToken to stop long running workbook.CalculateFormula in .NET | Aspose.Cells formula calculation timeout handling example | interrupt Excel formula evaluation with Aspose.Cells and CancellationToken | run Workbook.CalculateFormula on a background task with timeout in C#
// Tags: Aspose.Cells Workbook.CalculateFormula cancellation | c# cancellation token for Excel formula evaluation | timeout enforcement for workbook.CalculateFormula | background task execution of Aspose.Cells calculations | saving workbook after interrupted calculation

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, creates a CancellationTokenSource set to a 5‑second timeout, runs Workbook.CalculateFormula on a background task, waits for completion within the timeout, cancels the operation if it exceeds the limit, handles cancellation and other exceptions, and finally saves the workbook.
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

        Workbook workbook = null;
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Define a timeout for the calculation (5 seconds)
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

        try
        {
            // Run calculation on a background task so we can enforce the timeout
            Task calcTask = Task.Run(() => workbook.CalculateFormula(), cts.Token);

            // Wait for the task to complete within the timeout period
            bool completed = calcTask.Wait(TimeSpan.FromSeconds(5), cts.Token);
            if (!completed)
            {
                // Timeout elapsed – cancel the operation
                cts.Cancel();
                throw new OperationCanceledException(cts.Token);
            }

            Console.WriteLine("Formula calculation completed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Formula calculation was cancelled after the timeout period.");
        }
        catch (AggregateException ae)
        {
            // Unwrap any exceptions thrown inside the task
            foreach (var inner in ae.InnerExceptions)
            {
                Console.WriteLine($"Error during calculation: {inner.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }

        // Save the workbook (full or partial calculation)
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
