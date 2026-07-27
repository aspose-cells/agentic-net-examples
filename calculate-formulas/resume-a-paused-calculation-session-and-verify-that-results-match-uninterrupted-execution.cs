// Title: Resume an interrupted Aspose.Cells calculation and verify identical results (C#)
// Description: Demonstrates how to pause a long‑running CalculateFormula operation with ThreadInterruptMonitor, catch the interruption, resume the calculation, and compare the resumed cell values with those from an uninterrupted run to ensure they match.
// Keywords: Aspose.Cells resume calculation | ThreadInterruptMonitor C# | interrupt and continue formula evaluation | compare formula results Aspose.Cells | C# workbook calculation pause | validate resumed calculation
// Common Searches: how to resume Aspose.Cells calculation after interruption | Aspose.Cells ThreadInterruptMonitor example | compare interrupted vs full calculation results C# | resume workbook formula evaluation Aspose.Cells | verify calculation consistency after pause
// Developer Intent: Pause a formula calculation, resume it later, and confirm that the final values are the same as a full, uninterrupted calculation.
// Use Cases: Implement a cancellable UI for large spreadsheet calculations that can be resumed without data loss. | Create automated tests that ensure interrupt‑resume logic does not alter formula outcomes. | Process massive worksheets in background threads, allowing graceful interruption and later continuation.
// AI Prompts: Generate a C# snippet that uses Aspose.Cells ThreadInterruptMonitor to interrupt CalculateFormula after a set time, then resumes the calculation and returns true if all cell values match the original run. | Explain how to capture cell values before interruption and compare them after resuming, handling numeric and string type differences. | Suggest a unit‑test structure for verifying that an interrupted Aspose.Cells calculation produces identical results to an uninterrupted execution.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to pause a long‑running CalculateFormula operation with ThreadInterruptMonitor, catch the interruption, resume the calculation, and compare the resumed cell values with those from an uninterrupted run to ensure they match.
class ResumeCalculationDemo
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a workbook with sample data and formulas
        // -------------------------------------------------
        Workbook wbFull = new Workbook();
        Worksheet wsFull = wbFull.Worksheets[0];

        // Populate cells A and B, and set formula in C
        for (int i = 0; i < 100; i++)
        {
            wsFull.Cells[i, 0].PutValue(i);          // A column
            wsFull.Cells[i, 1].PutValue(i * 2);      // B column
            wsFull.Cells[i, 2].Formula = $"=A{i}+B{i}"; // C column = A+B
        }

        // -------------------------------------------------
        // 2. Perform uninterrupted calculation and store results
        // -------------------------------------------------
        wbFull.CalculateFormula();

        var fullResults = new List<object>();
        for (int i = 0; i < 100; i++)
        {
            fullResults.Add(wsFull.Cells[i, 2].Value);
        }

        // Save the workbook (used later for loading a fresh copy)
        wbFull.Save("FullCalc.xlsx");

        // -------------------------------------------------
        // 3. Load a fresh copy and interrupt the calculation
        // -------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        Workbook wbInterrupted = new Workbook("FullCalc.xlsx", loadOptions);

        // Set up a ThreadInterruptMonitor to request interruption quickly
        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(false);
        loadOptions.InterruptMonitor = monitor;
        monitor.StartMonitor(10); // 10 ms limit forces an early interruption

        try
        {
            wbInterrupted.CalculateFormula(); // This will be interrupted
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Calculation was interrupted as expected.");
        }
        finally
        {
            // Ensure the monitor thread is cleaned up
            monitor.FinishMonitor();
        }

        // -------------------------------------------------
        // 4. Resume calculation without any monitor
        // -------------------------------------------------
        wbInterrupted.CalculateFormula();

        // -------------------------------------------------
        // 5. Verify that resumed results match the uninterrupted ones
        // -------------------------------------------------
        bool match = true;
        Worksheet wsInterrupted = wbInterrupted.Worksheets[0];

        for (int i = 0; i < 100; i++)
        {
            object resumedValue = wsInterrupted.Cells[i, 2].Value;
            object expectedValue = fullResults[i];

            if (!object.Equals(resumedValue, expectedValue))
            {
                match = false;
                Console.WriteLine($"Mismatch at row {i}: expected {expectedValue}, got {resumedValue}");
                break;
            }
        }

        Console.WriteLine(match
            ? "Resumed calculation matches uninterrupted results."
            : "Results differ after resuming.");

        // Save the workbook after successful resume
        wbInterrupted.Save("ResumedCalc.xlsx");
    }
}
