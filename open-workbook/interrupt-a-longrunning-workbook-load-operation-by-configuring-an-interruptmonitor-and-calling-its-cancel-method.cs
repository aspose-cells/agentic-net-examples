using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Create an interrupt monitor that will throw an exception when interrupted
            ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(terminateWithoutException: false);

            // Configure load options to use the monitor
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Start monitoring with a 1‑second time limit
            monitor.StartMonitor(1000); // 1000 ms

            try
            {
                // Attempt to load a large workbook; this operation will be interrupted
                Workbook workbook = new Workbook("LargeWorkbook.xlsx", loadOptions);

                // If loading completes before the timeout, finish the monitor for this step
                monitor.FinishMonitor();

                // (Optional) Save the workbook if needed
                // workbook.Save("Result.xlsx");
                Console.WriteLine("Workbook loaded successfully (unexpected).");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Expected path: the load operation was interrupted
                Console.WriteLine("Load operation was successfully interrupted.");
            }
            catch (Exception ex)
            {
                // Any other unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}