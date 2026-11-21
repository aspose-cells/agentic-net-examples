// Title: How to cancel a long‑running Excel workbook load with Aspose.Cells InterruptMonitor in C#
// Common Searches: c# cancel Aspose.Cells workbook load using InterruptMonitor | stop loading large XLSX file with Aspose.Cells interrupt monitor | Aspose.Cells LoadOptions interrupt monitor example for .NET | cancel long running workbook loading thread Aspose.Cells | handle OperationCanceledException when aborting Aspose.Cells load
// Tags: InterruptMonitor usage with LoadOptions | cancelling workbook load Aspose.Cells | threaded Excel load cancellation .NET | OperationCanceledException handling Aspose.Cells | large XLSX load abort Aspose.Cells

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// Demonstrates loading a large XLSX workbook on a separate thread with LoadOptions that reference an InterruptMonitor, then issuing monitor.Interrupt() after a short delay to cancel the operation. The code catches OperationCanceledException, reports cancellation, and only saves the workbook if the load completes successfully.
class Program
{
    static void Main()
    {
        // Path to a large workbook that takes noticeable time to load
        string filePath = @"C:\Data\LargeWorkbook.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Source file not found: {filePath}");
            return;
        }

        // Create an InterruptMonitor instance
        InterruptMonitor monitor = new InterruptMonitor();

        // Configure LoadOptions to use the monitor
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            InterruptMonitor = monitor
        };

        // Variable to hold the workbook reference
        Workbook workbook = null;

        // Start the load operation on a separate thread
        Thread loadThread = new Thread(() =>
        {
            try
            {
                // This call will respect the InterruptMonitor
                workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Workbook load was cancelled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Load failed: {ex.Message}");
            }
        });

        loadThread.Start();

        // Wait a short period before cancelling (simulating user interruption)
        Thread.Sleep(2000); // 2 seconds

        // Request cancellation of the long‑running load operation
        monitor.Interrupt(); // Correct method to signal cancellation
        Console.WriteLine("Cancel request sent to InterruptMonitor.");

        // Wait for the load thread to finish
        loadThread.Join();

        // Proceed only if the workbook was loaded successfully
        if (workbook != null)
        {
            try
            {
                string outputPath = @"C:\Data\ProcessedWorkbook.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save to a new file (demonstrating normal lifecycle)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
