// Title: Cancel a long‑running workbook load in Aspose.Cells for .NET with InterruptMonitor
// Common Searches: c# cancel workbook load using Aspose.Cells InterruptMonitor | how to interrupt a large XLSX load in Aspose.Cells .NET | example of stopping Aspose.Cells workbook loading with Task and Interrupt() | aspose.cells loadoptions interruptmonitor usage to abort loading | timeout cancellation of Excel workbook load in C# with Aspose.Cells
// Tags: interruptmonitor workbook load cancellation c# | aspose.cells loadoptions interruptmonitor integration | c# task based workbook loading with cancellation | aspose.cells cancel long running excel load | handling interrupt exception Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

// The example creates an InterruptMonitor, assigns it to LoadOptions, and loads a large workbook on a background Task. After a short delay it calls monitor.Interrupt() to signal cancellation; the load operation checks the monitor periodically, throws an exception, and the task can be awaited with proper cancellation handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create an interrupt monitor that can cancel long‑running operations
            InterruptMonitor monitor = new InterruptMonitor();

            // Configure load options to use the interrupt monitor
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                InterruptMonitor = monitor
            };

            // Path to a large workbook that takes time to load
            string filePath = "LargeWorkbook.xlsx";

            // Verify that the workbook file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File \"{filePath}\" not found.");
                return;
            }

            // Start loading the workbook on a separate thread so we can cancel it
            Task loadTask = Task.Run(() =>
            {
                try
                {
                    // The load operation will periodically check the monitor
                    Workbook wb = new Workbook(filePath, loadOptions);
                    // Loading completed successfully; you can work with the workbook here
                    Console.WriteLine("Workbook loaded successfully.");
                }
                catch (Exception ex)
                {
                    // Handle loading errors or cancellation
                    Console.WriteLine($"Workbook loading failed: {ex.Message}");
                    throw;
                }
            });

            // Wait briefly before issuing the cancel request
            Thread.Sleep(2000); // e.g., wait 2 seconds

            // Request cancellation of the loading operation
            // The correct method to signal interruption is Interrupt()
            monitor.Interrupt();

            // Wait for the load task to finish and handle cancellation
            try
            {
                loadTask.Wait();
            }
            catch (AggregateException ae)
            {
                // Loading was cancelled or failed; handle as needed
                foreach (var inner in ae.InnerExceptions)
                {
                    Console.WriteLine($"Task error: {inner.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Global exception handling
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
