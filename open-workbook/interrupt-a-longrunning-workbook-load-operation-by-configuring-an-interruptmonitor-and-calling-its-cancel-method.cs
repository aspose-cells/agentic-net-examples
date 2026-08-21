// Title: Cancel Long‑Running Workbook Load with ThreadInterruptMonitor in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to configure a ThreadInterruptMonitor, attach it to LoadOptions, start a 500 ms timeout, and abort a Workbook load that exceeds the limit. The sample catches the CellsException with ExceptionType.Interrupted and shows proper monitor cleanup on success.
// Keywords: Aspose.Cells | ThreadInterruptMonitor | interrupt workbook load | cancel long running load | LoadOptions InterruptMonitor | C# Excel timeout | CellsException Interrupted | .NET Excel performance | abort workbook loading
// Common Searches: Aspose.Cells interrupt workbook load C# | ThreadInterruptMonitor timeout example | How to cancel Excel file loading in .NET | Catch CellsException Interrupted Aspose | Set load timeout for large Excel file
// Developer Intent: Implement a time‑bound workbook loading process that automatically stops if it exceeds a predefined duration, and handle the interruption gracefully.
// Use Cases: Prevent UI freeze by limiting Excel load time to 500 ms or less. | Safely abort loading of massive spreadsheets in server‑side processing. | Log and respond to load interruptions without crashing the application. | Release monitor resources after a successful load to avoid memory leaks.
// AI Prompts: Write C# code that uses ThreadInterruptMonitor to abort a Workbook load after 1 second and logs the interruption event. | Show how to wrap Aspose.Cells Workbook loading in a try‑catch that handles CellsException.Interrupted and ensures the monitor is finished on success.

using System;
using Aspose.Cells;

// Demonstrates how to configure a ThreadInterruptMonitor, attach it to LoadOptions, start a 500 ms timeout, and abort a Workbook load that exceeds the limit. The sample catches the CellsException with ExceptionType.Interrupted and shows proper monitor cleanup on success.
class InterruptLoadDemo
{
    static void Main()
    {
        // Create a ThreadInterruptMonitor that will request interruption after a time limit.
        // The boolean parameter indicates whether to terminate silently (false = throw exception).
        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(false);

        // Configure LoadOptions to use the interrupt monitor.
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        // Start monitoring with a 500 ms time limit.
        // If the load operation exceeds this limit, the monitor will request interruption.
        monitor.StartMonitor(500);

        try
        {
            // Attempt to load a workbook that may take a long time.
            // The load will be interrupted if the time limit is exceeded.
            Workbook workbook = new Workbook("LargeWorkbook.xlsx", loadOptions);

            // If loading completes before interruption, finish the monitor for this procedure.
            monitor.FinishMonitor();

            // Optional: save the workbook to verify successful load.
            workbook.Save("Result.xlsx");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            // Expected path when the load operation is interrupted.
            Console.WriteLine("Loading operation was interrupted.");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
