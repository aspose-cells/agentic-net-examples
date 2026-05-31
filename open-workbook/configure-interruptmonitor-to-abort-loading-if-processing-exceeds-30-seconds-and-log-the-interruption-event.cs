using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Create an interrupt monitor that will throw an exception when interrupted
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(terminateWithoutException: false);

            // Configure load options to use the monitor
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Start monitoring with a 30‑second (30000 ms) time limit
            monitor.StartMonitor(30000);

            try
            {
                // Attempt to load the workbook using the configured load options
                Workbook workbook = new Workbook("LargeFile.xlsx", loadOptions);

                // Optional: save the workbook after successful load
                workbook.Save("LoadedResult.xlsx");
                Console.WriteLine("Workbook loaded and saved successfully.");
            }
            catch (Exception ex)
            {
                // Log the interruption event
                Console.WriteLine("Loading was interrupted: " + ex.Message);
            }
        }
    }
}