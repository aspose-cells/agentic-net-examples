// Title: Pause Excel formula calculation after a time threshold and resume it later with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells pause formula calculation after 15 seconds | resume Excel workbook calculation after timeout using Aspose.Cells | save intermediate workbook state during long formula evaluation Aspose.Cells | how to use CancellationToken with Aspose.Cells CalculateFormula | run Aspose.Cells CalculateFormula with time limit and continue later
// Tags: Aspose.Cells calculate formula with timeout | pause and resume workbook calculation C# | save intermediate workbook state Aspose.Cells | cancellation token for Excel formula calculation | time‑bounded Excel calculation Aspose.Cells

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file with Aspose.Cells, starts formula calculation on a background task, and waits for a configurable time limit. If the calculation exceeds the limit, the task is cancelled, the workbook is saved to a temporary file, then reloaded and the full calculation is performed before saving the final result.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "Input.xlsx";
                const string pausedPath = "Workbook_Paused.xlsx";
                const string finalPath = "Workbook_Final.xlsx";

                // Verify the input file exists before loading.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                int timeThresholdSeconds = 15;
                using CancellationTokenSource cts = new CancellationTokenSource();

                // Run formula calculation in a separate task.
                Task calculationTask = Task.Run(() =>
                {
                    workbook.CalculateFormula();
                }, cts.Token);

                // Wait for completion or timeout.
                bool completedInTime = calculationTask.Wait(TimeSpan.FromSeconds(timeThresholdSeconds));

                if (!completedInTime)
                {
                    // Cancel the task (if it respects cancellation) and save current state.
                    cts.Cancel();
                    workbook.Save(pausedPath);
                    Console.WriteLine($"Calculation paused. Workbook saved to {pausedPath}");

                    // Load the paused workbook and resume calculation.
                    if (!File.Exists(pausedPath))
                    {
                        Console.WriteLine($"Paused workbook not found: {pausedPath}");
                        return;
                    }

                    Workbook resumedWorkbook = new Workbook(pausedPath);

                    try
                    {
                        // Perform full calculation on the resumed workbook.
                        resumedWorkbook.CalculateFormula();

                        // Save the fully calculated workbook.
                        resumedWorkbook.Save(finalPath);
                        Console.WriteLine($"Final workbook saved to {finalPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during resumed calculation: {ex.Message}");
                    }
                }
                else
                {
                    // Calculation completed within the time threshold.
                    workbook.Save(finalPath);
                    Console.WriteLine($"Calculation completed in time. Workbook saved to {finalPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
