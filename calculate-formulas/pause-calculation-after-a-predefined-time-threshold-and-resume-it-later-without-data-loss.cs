// Title: Pause and Resume Formula Calculation with SystemTimeInterruptMonitor in Aspose.Cells for C#/.NET
// Description: Shows how to set a time limit on workbook formula evaluation, capture the interruption, and later continue the computation without losing any intermediate values. The sample creates a 10,000‑row sheet, forces a 500 ms timeout, then restarts with a longer limit and writes the fully calculated workbook to disk.
// Keywords: Aspose.Cells C# | .NET formula calculation timeout | SystemTimeInterruptMonitor usage | interrupt and continue spreadsheet calculation | large workbook performance | resume interrupted calculation | UI‑responsive Excel processing | server‑side spreadsheet quota handling
// Common Searches: Aspose.Cells limit formula calculation time | How to continue a stopped calculation in Aspose.Cells | SystemTimeInterruptMonitor example C# | Resume workbook calculation after timeout | Break large spreadsheet calculation into chunks
// Developer Intent: The developer needs to stop a long‑running formula evaluation after a predefined duration and then pick up the remaining work later, ensuring no data is lost.
// Use Cases: Split heavy spreadsheet calculations into timed segments to keep a desktop UI responsive. | Honor execution‑time quotas in cloud services by pausing calculations when the limit is reached and finishing them in a subsequent request. | Defer resource‑intensive formula processing to off‑peak hours while preserving already computed results.
// AI Prompts: Generate C# code that uses SystemTimeInterruptMonitor to halt a calculation after 1 second and then resume it with a new timeout. | Explain how to detect which cells remain unevaluated after an interruption and log their addresses before resuming. | Provide guidance on handling different CellsException codes when working with InterruptMonitor in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to set a time limit on workbook formula evaluation, capture the interruption, and later continue the computation without losing any intermediate values. The sample creates a 10,000‑row sheet, forces a 500 ms timeout, then restarts with a longer limit and writes the fully calculated workbook to disk.
class PauseResumeCalculationDemo
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a workbook and populate it with data and formulas
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill 10,000 rows with a value and a simple formula
        for (int i = 0; i < 10000; i++)
        {
            sheet.Cells[i, 0].PutValue(i);                     // Column A: raw value
            sheet.Cells[i, 1].Formula = $"=A{i}+10";          // Column B: formula based on column A
        }

        // -------------------------------------------------
        // 2. Attach a SystemTimeInterruptMonitor to the workbook
        // -------------------------------------------------
        // terminateWithoutException = false -> an exception will be thrown when time limit is exceeded
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);
        workbook.InterruptMonitor = monitor;

        // -------------------------------------------------
        // 3. First calculation attempt with a short time limit (pause scenario)
        // -------------------------------------------------
        monitor.StartMonitor(500); // 500 ms time limit

        try
        {
            workbook.CalculateFormula(); // Start calculation
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Calculation paused: time limit reached.");
        }

        // -------------------------------------------------
        // 4. Resume calculation with a longer time limit
        // -------------------------------------------------
        monitor.StartMonitor(2000); // 2 seconds time limit for the remaining work

        try
        {
            workbook.CalculateFormula(); // Continue calculation from where it stopped
            Console.WriteLine("Calculation completed after resume.");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Unexpected interruption: {ex.Message}");
        }

        // -------------------------------------------------
        // 5. Save the workbook – data is intact and fully calculated
        // -------------------------------------------------
        workbook.Save("PausedResumeResult.xlsx");
    }
}
