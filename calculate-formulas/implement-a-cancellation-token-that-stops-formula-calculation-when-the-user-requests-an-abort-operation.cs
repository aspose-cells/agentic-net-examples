// Title: Use a CancellationToken to abort Aspose.Cells workbook.CalculateFormula in a C# console app
// Common Searches: how to stop Aspose.Cells CalculateFormula execution using CancellationToken in .NET | cancel long‑running Excel formula calculation with Aspose.Cells C# console | abort workbook.CalculateFormula when user presses a key in C# | implement user‑initiated cancellation for Aspose.Cells formula evaluation | handle OperationCanceledException during Aspose.Cells CalculateFormula
// Tags: Aspose.Cells CalculateFormula cancellation support | C# console workbook calculation abort | user‑initiated formula evaluation stop | OperationCanceledException handling Aspose.Cells | background task formula calculation Aspose.Cells

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads an Excel workbook (creating a new one if the file is missing), runs workbook.CalculateFormula on a background task, and monitors console input for the 'c' key. A CancellationTokenSource lets the user request cancellation, which aborts the calculation task and is caught as an OperationCanceledException. After cancellation or successful completion, the workbook is saved to the output file.
class FormulaCalculationWithCancellation
{
    static void Main()
    {
        // Path to the input and output files.
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Load the workbook, creating a new one if the input file does not exist.
        Workbook workbook;
        try
        {
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Set up a cancellation token source that can be triggered by the user.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            CancellationToken token = cts.Token;

            // Start the calculation on a separate task so we can monitor user input.
            Task calculationTask = Task.Run(() =>
            {
                // Perform the calculation. The overload does not accept a token,
                // so we simply run it; the token can cancel the task before it starts.
                workbook.CalculateFormula();
            }, token);

            Console.WriteLine("Press 'c' to cancel formula calculation...");

            // Simple loop waiting for user to press 'c' to request cancellation.
            while (!calculationTask.IsCompleted)
            {
                if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.C)
                {
                    // Signal cancellation.
                    cts.Cancel();
                    Console.WriteLine("Cancellation requested.");
                    break;
                }

                // Small delay to avoid busy‑waiting.
                Thread.Sleep(100);
            }

            try
            {
                // Wait for the calculation task to finish (or be cancelled).
                calculationTask.Wait(token);
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Formula calculation was cancelled.");
            }
            catch (AggregateException ae)
            {
                // Re‑throw unexpected exceptions.
                foreach (var ex in ae.InnerExceptions)
                {
                    Console.WriteLine($"Error during calculation: {ex.Message}");
                }
                return;
            }

            // Save the workbook (if calculation completed) using the save rule.
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
}
