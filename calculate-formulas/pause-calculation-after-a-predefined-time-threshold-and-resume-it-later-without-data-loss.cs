// Title: How to pause and later resume a long‑running workbook formula calculation with a time limit using Aspose.Cells for .NET
// AI Prompts: Demonstrate using SystemTimeInterruptMonitor to interrupt Workbook.CalculateFormula after a specified number of milliseconds and then continue the calculation from the same state. | Provide C# code that catches the CellsException thrown on interruption, restarts the monitor with a new timeout, and finishes the remaining formulas without losing previously computed results.
// Common Searches: c# Aspose.Cells interrupt CalculateFormula after 1 second and resume later | SystemTimeInterruptMonitor example for limiting formula calculation time | how to handle CellsException.Interrupted during large workbook calculation | resume workbook calculation after timeout using Aspose.Cells .NET | preserve partial results when pausing Aspose.Cells formula evaluation
// Tags: SystemTimeInterruptMonitor timeout handling | interrupt Aspose.Cells formula calculation | resume workbook.CalculateFormula after interruption | partial calculation preservation Aspose.Cells | C# large workbook formula evaluation performance

using System;
using Aspose.Cells;

// The example creates a workbook filled with many SUM formulas, attaches a SystemTimeInterruptMonitor, runs CalculateFormula with a 1‑second limit causing an interruption, then restarts the monitor with a 2‑second limit to finish the remaining calculations, and finally saves the workbook while preserving any data computed before the interruptions.
class PauseResumeCalculation
{
    static void Main()
    {
        // Create a new workbook and populate it with many formulas to make calculation time-consuming
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        for (int row = 0; row < 5000; row++)
        {
            // Each cell contains a formula that references a range; this creates a heavy calculation load
            sheet.Cells[row, 0].Formula = $"=SUM(A{row + 1}:A{row + 10})";
        }

        // Attach a SystemTimeInterruptMonitor to the workbook.
        // terminateWithoutException = false means an exception will be thrown when the time limit is reached.
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);
        workbook.InterruptMonitor = monitor;

        // ---------- First calculation attempt (pause after 1 second) ----------
        monitor.StartMonitor(1000); // time limit in milliseconds

        try
        {
            workbook.CalculateFormula();
            Console.WriteLine("Calculation finished within the first time slice.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Calculation paused after reaching the first time limit.");
        }

        // ---------- Resume calculation (allow another 2 seconds) ----------
        monitor.StartMonitor(2000); // new time limit

        try
        {
            workbook.CalculateFormula();
            Console.WriteLine("Remaining calculation completed successfully.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Calculation paused again after reaching the second time limit.");
        }

        // Save the workbook; data generated before interruption is preserved.
        workbook.Save("Result.xlsx");
    }
}
