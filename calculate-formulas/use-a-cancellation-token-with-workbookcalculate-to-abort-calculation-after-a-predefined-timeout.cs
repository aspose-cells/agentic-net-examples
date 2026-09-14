// Title: How to abort Aspose.Cells Workbook.CalculateFormula with a timeout using ThreadInterruptMonitor in C#
// AI Prompts: Generate C# code that configures a ThreadInterruptMonitor with a 2‑second limit, runs Workbook.CalculateFormula, and catches the interruption exception. | Show how to handle a CellsException of type Interrupted after a calculation timeout and continue program execution. | Provide an example of saving the workbook when the calculation is stopped early due to the timeout.
// Common Searches: Aspose.Cells stop CalculateFormula after 2000 ms in C# | C# ThreadInterruptMonitor example for aborting long Excel calculations | How to handle interrupted calculation exception in Aspose.Cells workbook
// Tags: Aspose.Cells calculation timeout ThreadInterruptMonitor | Workbook.CalculateFormula interruption handling | catch CellsException Interrupted Aspose.Cells | save partially calculated workbook Aspose.Cells | C# Excel formula evaluation timeout

using System;
using Aspose.Cells;

// The sample creates a workbook with thousands of simple formulas, attaches a ThreadInterruptMonitor set to a 2000 ms limit, invokes Workbook.CalculateFormula, catches the CellsException when the calculation exceeds the timeout, stops the monitor, and saves the workbook containing whatever results were computed before the interruption.
class WorkbookCalculationWithTimeout
{
    static void Main()
    {
        // Create a new workbook and add sample data/formulas
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate many rows to make calculation take noticeable time
        for (int i = 0; i < 5000; i++)
        {
            sheet.Cells[i, 0].PutValue(i);                     // Column A values
            sheet.Cells[i, 1].Formula = $"=A{i}+1";           // Column B formulas
        }

        // Create a thread‑based interrupt monitor
        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(false);
        // Assign the monitor to the workbook
        workbook.InterruptMonitor = monitor;

        // Start the monitor with a time limit (e.g., 2000 ms = 2 seconds)
        monitor.StartMonitor(2000);

        try
        {
            // Perform calculation; it will be interrupted if it exceeds the time limit
            workbook.CalculateFormula();
            Console.WriteLine("Calculation completed successfully.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            // Handle the interruption caused by the timeout
            Console.WriteLine("Calculation was interrupted due to timeout.");
        }
        finally
        {
            // Ensure the monitor is stopped
            monitor.FinishMonitor();
        }

        // Save the workbook (optional, will contain whatever was calculated before interruption)
        workbook.Save("Result.xlsx");
    }
}
