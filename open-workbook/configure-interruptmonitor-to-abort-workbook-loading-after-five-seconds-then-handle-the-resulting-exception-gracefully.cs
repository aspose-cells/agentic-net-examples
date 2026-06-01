using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Create an interrupt monitor that throws an exception when interrupted
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(terminateWithoutException: false);

            // Assign the monitor to LoadOptions
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Start monitoring with a 5‑second (5000 ms) time limit
            monitor.StartMonitor(5000);

            try
            {
                // Attempt to load a workbook; will be aborted if loading exceeds 5 seconds
                Workbook workbook = new Workbook("large.xlsx", loadOptions);
                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (Exception ex)
            {
                // Gracefully handle the interruption exception
                Console.WriteLine("Loading aborted: " + ex.Message);
            }
        }
    }
}