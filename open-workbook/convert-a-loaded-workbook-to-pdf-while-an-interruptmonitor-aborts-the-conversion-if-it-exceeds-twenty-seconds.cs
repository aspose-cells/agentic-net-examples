using System;
using Aspose.Cells;

class ConvertToPdfWithInterrupt
{
    static void Main()
    {
        // Load the source workbook (replace with actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create an interrupt monitor that will throw an exception when interrupted
        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(terminateWithoutException: false);

        // Assign the monitor to the workbook
        workbook.InterruptMonitor = monitor;

        // Start monitoring with a 20‑second time limit (20000 ms)
        monitor.StartMonitor(20000);

        try
        {
            // Attempt to save the workbook as PDF; will be aborted if it exceeds the time limit
            workbook.Save("output.pdf", SaveFormat.Pdf);
            Console.WriteLine("Workbook successfully converted to PDF.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            // Handle the interruption caused by the time limit
            Console.WriteLine("Conversion was interrupted because it exceeded the 20‑second limit.");
        }
        finally
        {
            // Clean up the monitor after the operation
            monitor.FinishMonitor();
        }
    }
}