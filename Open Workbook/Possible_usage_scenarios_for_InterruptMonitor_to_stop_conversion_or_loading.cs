using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace InterruptMonitorDemo
{
    class Program
    {
        static void Main()
        {
            // Scenario 1: Interrupt loading of a workbook using a time‑based monitor.
            InterruptLoadWithTimeLimit();

            // Scenario 2: Interrupt a conversion operation (e.g., save to PDF) using Interrupt().
            InterruptConversionDemo();
        }

        static void InterruptLoadWithTimeLimit()
        {
            // Create a ThreadInterruptMonitor that throws an exception when interrupted.
            ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(terminateWithoutException: false);

            // Attach the monitor to LoadOptions.
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Start monitoring – allow only 1000 ms for the load operation.
            monitor.StartMonitor(1000);

            try
            {
                // Load a workbook with the specified LoadOptions.
                Workbook wb = new Workbook("LargeFile.xlsx", loadOptions);

                // Loading finished before timeout; stop the monitor.
                monitor.FinishMonitor();

                Console.WriteLine("Workbook loaded successfully within the time limit.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // The load operation exceeded the time limit and was interrupted.
                Console.WriteLine("Loading was interrupted due to time limit.");
            }
            finally
            {
                // Ensure the monitor thread is cleaned up even if an exception occurs.
                monitor.FinishMonitor();
            }
        }

        static void InterruptConversionDemo()
        {
            // Create a new workbook and add some sample data.
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Sample data for PDF conversion.");

            // Create a simple InterruptMonitor and assign it to the workbook.
            InterruptMonitor monitor = new InterruptMonitor();
            wb.InterruptMonitor = monitor;

            // In a background task, request interruption after a short delay.
            Task.Run(() =>
            {
                Thread.Sleep(500); // Wait 500 ms.
                monitor.Interrupt(); // Signal interruption.
            });

            try
            {
                // Attempt to save the workbook as PDF; this should be interrupted.
                wb.Save("Converted.pdf", SaveFormat.Pdf);
                Console.WriteLine("PDF conversion completed.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // The save operation was interrupted by the monitor.
                Console.WriteLine("PDF conversion was interrupted by the monitor.");
            }
        }
    }
}