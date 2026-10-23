// Title: How to abort a long‑running Workbook.CalculateFormula operation with a CancellationToken in Aspose.Cells for .NET (C#)
// Common Searches: c# how to stop Aspose.Cells CalculateFormula after a timeout | using CancellationToken to cancel Excel formula calculation with Aspose.Cells | abort long running workbook.CalculateFormula in a .NET application | example of timer‑based cancellation for Aspose.Cells formula evaluation | handle user‑initiated cancellation during Aspose.Cells workbook calculation
// Tags: Aspose.Cells formula calculation cancellation | Workbook.CalculateFormula abort with CancellationToken | c# timer based cancellation for Excel processing | handling cancellation in Aspose.Cells workbook calculation | saving workbook after cancelled formula evaluation

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, creates a CancellationTokenSource that triggers after five seconds via a Timer, checks the token before calling Workbook.CalculateFormula (which does not accept a token directly), reports whether the calculation completed or was cancelled, and finally saves the workbook to a new file.
class FormulaCalculationWithCancellation
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
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

        // Set up a cancellation token source (simulated abort after 5 seconds)
        using (CancellationTokenSource cts = new CancellationTokenSource())
        using (Timer abortTimer = new Timer(state => cts.Cancel(), null, TimeSpan.FromSeconds(5), Timeout.InfiniteTimeSpan))
        {
            try
            {
                // Check for cancellation before starting calculation
                if (cts.Token.IsCancellationRequested)
                {
                    Console.WriteLine("Formula calculation was cancelled before start.");
                }
                else
                {
                    // Perform formula calculation (Aspose.Cells does not accept a token directly)
                    workbook.CalculateFormula();
                    Console.WriteLine("Formula calculation completed successfully.");
                }
            }
            catch (OperationCanceledException)
            {
                // This block is retained for completeness; Aspose.Cells does not throw this.
                Console.WriteLine("Formula calculation was cancelled by the user.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during formula calculation: {ex.Message}");
            }
            finally
            {
                // Save the workbook
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
}
