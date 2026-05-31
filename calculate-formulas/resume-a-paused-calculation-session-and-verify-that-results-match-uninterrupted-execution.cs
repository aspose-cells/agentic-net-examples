using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCalculationResumeDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for temporary files
            string originalPath = "OriginalWorkbook.xlsx";
            string fullCalcPath = "FullCalculation.xlsx";
            string resumedPath = "ResumedCalculation.xlsx";

            try
            {
                // -------------------------------------------------
                // 1. Create a workbook with many formulas
                // -------------------------------------------------
                Workbook wbOriginal = new Workbook();
                Worksheet ws = wbOriginal.Worksheets[0];

                // Fill data and formulas in a 200x10 range
                for (int row = 0; row < 200; row++)
                {
                    ws.Cells[row, 0].PutValue(row + 1);                     // Column A: simple numbers
                    ws.Cells[row, 1].PutValue((row + 1) * 2);               // Column B: simple numbers
                    ws.Cells[row, 2].Formula = $"=A{row + 1}+B{row + 1}";   // Column C: A+B
                    ws.Cells[row, 3].Formula = $"=SUM(A{row + 1}:C{row + 1})"; // Column D: sum A‑C
                }

                // Save the original workbook (contains formulas only)
                try
                {
                    wbOriginal.Save(originalPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save original workbook: {ex.Message}");
                    return;
                }

                // -------------------------------------------------
                // 2. Perform uninterrupted full calculation and store results
                // -------------------------------------------------
                if (!File.Exists(originalPath))
                {
                    Console.WriteLine($"File not found: {originalPath}");
                    return;
                }

                Workbook wbFull;
                try
                {
                    wbFull = new Workbook(originalPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load original workbook for full calculation: {ex.Message}");
                    return;
                }

                wbFull.CalculateFormula(); // Full calculation without interruption

                try
                {
                    wbFull.Save(fullCalcPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save full calculation workbook: {ex.Message}");
                    return;
                }

                // Capture results for later verification
                double[,] fullResults = new double[200, 4];
                Worksheet wsFull = wbFull.Worksheets[0];
                for (int r = 0; r < 200; r++)
                {
                    for (int c = 0; c < 4; c++)
                    {
                        fullResults[r, c] = wsFull.Cells[r, c].DoubleValue;
                    }
                }

                // -------------------------------------------------
                // 3. Simulate a paused calculation using ThreadInterruptMonitor
                // -------------------------------------------------
                if (!File.Exists(originalPath))
                {
                    Console.WriteLine($"File not found: {originalPath}");
                    return;
                }

                Workbook wbPaused;
                try
                {
                    wbPaused = new Workbook(originalPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load original workbook for paused calculation: {ex.Message}");
                    return;
                }

                // Create a monitor that will interrupt after a short time (e.g., 100 ms)
                ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(terminateWithoutException: false);
                LoadOptions loadOptions = new LoadOptions { InterruptMonitor = monitor };

                // Start monitoring before calculation
                try
                {
                    monitor.StartMonitor(100); // 100 ms limit to force interruption
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to start interrupt monitor: {ex.Message}");
                    return;
                }

                try
                {
                    // This call is expected to be interrupted and throw CellsException with code Interrupted
                    wbPaused.CalculateFormula();
                    Console.WriteLine("Unexpected: calculation completed without interruption.");
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
                {
                    Console.WriteLine("Calculation was successfully interrupted.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected exception during interruption: {ex.Message}");
                }
                finally
                {
                    // Ensure the monitor thread is finished
                    monitor.FinishMonitor();
                }

                // -------------------------------------------------
                // 4. Resume the calculation after interruption
                // -------------------------------------------------
                try
                {
                    monitor.StartMonitor(5000); // 5 seconds generous limit
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to restart interrupt monitor: {ex.Message}");
                    return;
                }

                try
                {
                    // Resume calculation from the point it was interrupted
                    wbPaused.CalculateFormula();
                    Console.WriteLine("Resumed calculation completed.");
                }
                catch (CellsException ex)
                {
                    Console.WriteLine($"Unexpected CellsException during resume: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected exception during resume: {ex.Message}");
                }
                finally
                {
                    monitor.FinishMonitor();
                }

                // Save the workbook after resumed calculation
                try
                {
                    wbPaused.Save(resumedPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save resumed workbook: {ex.Message}");
                    return;
                }

                // -------------------------------------------------
                // 5. Verify that resumed results match the uninterrupted results
                // -------------------------------------------------
                Worksheet wsResumed = wbPaused.Worksheets[0];
                bool allMatch = true;
                for (int r = 0; r < 200; r++)
                {
                    for (int c = 0; c < 4; c++)
                    {
                        double resumedValue = wsResumed.Cells[r, c].DoubleValue;
                        double fullValue = fullResults[r, c];
                        if (Math.Abs(resumedValue - fullValue) > 1e-9)
                        {
                            allMatch = false;
                            Console.WriteLine($"Mismatch at cell {CellsHelper.ColumnIndexToName(c)}{r + 1}: " +
                                              $"Resumed={resumedValue}, Full={fullValue}");
                        }
                    }
                }

                Console.WriteLine(allMatch
                    ? "Verification succeeded: resumed calculation matches uninterrupted calculation."
                    : "Verification failed: differences were found.");

                // -------------------------------------------------
                // Clean up temporary files (optional)
                // -------------------------------------------------
                // File.Delete(originalPath);
                // File.Delete(fullCalcPath);
                // File.Delete(resumedPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}