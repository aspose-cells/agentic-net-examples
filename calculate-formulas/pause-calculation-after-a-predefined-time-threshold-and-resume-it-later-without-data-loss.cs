// Title: Pause and Resume Formula Calculation with Aspose.Cells InterruptMonitor in C#
// Description: Demonstrates how to limit a long‑running workbook calculation using Aspose.Cells' SystemTimeInterruptMonitor, interrupt after a set time, and later resume without losing any computed results. The example creates 5,000 rows of formulas, triggers a 500 ms timeout, then continues with a larger limit and saves the workbook.
// Keywords: Aspose.Cells | InterruptMonitor | SystemTimeInterruptMonitor | pause calculation | resume calculation | C# | calculate formulas | timeout handling | incremental spreadsheet processing | .NET
// Common Searches: Aspose.Cells interrupt calculation after timeout | resume paused formula calculation C# | SystemTimeInterruptMonitor example | break long spreadsheet calculation into chunks | how to use CalculateFormula with time limit
// Developer Intent: Implement a time‑bounded formula calculation that can be halted and later continued, ensuring all formulas are eventually evaluated and the workbook remains intact.
// Use Cases: Keep a UI responsive by processing large workbooks in short time slices. | Enforce execution limits for server‑side spreadsheet tasks to avoid runaway processes. | Support retry or pause‑resume scenarios in background services that handle massive calculations.
// AI Prompts: Show C# code that uses Aspose.Cells SystemTimeInterruptMonitor to pause calculation after 500 ms and then resume. | Generate a try‑catch block that detects CellsException.Interrupted and restarts workbook.CalculateFormula with a new monitor. | Explain best practices for safely interrupting and resuming formula calculations in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Demonstrates how to limit a long‑running workbook calculation using Aspose.Cells' SystemTimeInterruptMonitor, interrupt after a set time, and later resume without losing any computed results. The example creates 5,000 rows of formulas, triggers a 500 ms timeout, then continues with a larger limit and saves the workbook.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and fill it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate 5000 rows with simple formulas to create a time‑consuming calculation
            for (int i = 0; i < 5000; i++)
            {
                sheet.Cells[i, 0].PutValue(i);                     // A column – numbers
                sheet.Cells[i, 1].Formula = $"=A{i + 1}*2";        // B column – formula
            }

            // 2. Create an interrupt monitor (throws exception when time limit is exceeded)
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(terminateWithoutException: false);
            workbook.InterruptMonitor = monitor;

            // 3. Start monitoring with a short time limit (e.g., 500 ms)
            monitor.StartMonitor(500);

            try
            {
                // Attempt to calculate all formulas; this will be interrupted after 500 ms
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed within the first time slice.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted after the time threshold.");
            }

            // 4. Resume calculation by starting a new monitor with a larger limit
            monitor.StartMonitor(2000); // allow up to 2 seconds for the remaining work

            try
            {
                // Continue the calculation from where it left off
                workbook.CalculateFormula();
                Console.WriteLine("Remaining calculation completed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // In a real scenario you could repeat the pause/resume cycle
                Console.WriteLine("Calculation was interrupted again.");
                throw; // rethrow or handle as needed
            }

            // 5. Save the workbook – data is intact because calculation finished
            workbook.Save("InterruptedCalculationResult.xlsx");
            Console.WriteLine("Workbook saved.");
        }
    }
}
