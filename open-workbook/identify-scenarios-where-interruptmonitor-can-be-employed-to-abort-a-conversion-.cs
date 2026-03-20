using System;
using System.Threading;
using Aspose.Cells;

class InterruptMonitorScenarios
{
    static void Main()
    {
        // Scenario 1: Abort a conversion (e.g., saving to PDF) using a manual interrupt request.
        AbortConversion();

        // Scenario 2: Abort a loading operation using a time‑based monitor.
        AbortLoading();
    }

    static void AbortConversion()
    {
        // Create a workbook and populate it with data to make the save operation noticeable.
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        for (int i = 0; i < 10000; i++)
        {
            ws.Cells[i, 0].PutValue($"Row {i}");
        }

        // Assign a simple InterruptMonitor to the workbook.
        InterruptMonitor monitor = new InterruptMonitor();
        wb.InterruptMonitor = monitor;

        // In a background thread request interruption after a short delay.
        ThreadPool.QueueUserWorkItem(_ =>
        {
            Thread.Sleep(500); // wait 500 ms
            monitor.Interrupt(); // request interruption
        });

        try
        {
            // Attempt to save the workbook as PDF; this should be interrupted.
            wb.Save("ConversionInterrupted.pdf", SaveFormat.Pdf);
            Console.WriteLine("Conversion completed (unexpected).");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Conversion was successfully interrupted.");
        }
    }

    static void AbortLoading()
    {
        // Path to a large workbook that may take noticeable time to load.
        string largeFilePath = "LargeWorkbook.xlsx";

        // Create a ThreadInterruptMonitor with exception termination (false).
        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(terminateWithoutException: false);
        LoadOptions loadOpts = new LoadOptions { InterruptMonitor = monitor };

        // Start monitoring with a 1‑second time limit.
        monitor.StartMonitor(1000); // 1000 ms

        try
        {
            // Load the workbook; if loading exceeds the limit, an interruption exception is thrown.
            Workbook wb = new Workbook(largeFilePath, loadOpts);
            Console.WriteLine("Workbook loaded successfully (unexpected).");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Loading was interrupted due to time limit.");
        }
        finally
        {
            // Clean up the monitor thread.
            monitor.FinishMonitor();
        }
    }
}