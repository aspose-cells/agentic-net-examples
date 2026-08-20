// Title: Cancel Long‑Running TIFF Export with Aspose.Cells InterruptMonitor and CancellationToken in C#
// Description: Demonstrates how to abort a time‑consuming TIFF conversion in Aspose.Cells by linking a CancellationToken to an InterruptMonitor, handling the CellsException Interrupted, and cleaning up resources.
// Keywords: Aspose.Cells TIFF cancellation | InterruptMonitor C# | CancellationToken Aspose.Cells | abort image export .NET | stop long running sheet render | CellsException Interrupted handling | C# workbook to TIFF cancel | Aspose.Cells rendering cancellation | user‑initiated abort TIFF conversion | aspose.cells cancel token example
// Common Searches: how to cancel TIFF export in Aspose.Cells | use CancellationToken with InterruptMonitor Aspose.Cells | abort long running sheet rendering C# | stop Aspose.Cells image conversion on user request | catch CellsException Interrupted Aspose.Cells
// Developer Intent: Implement a CancellationToken that triggers InterruptMonitor to stop a TIFF conversion in Aspose.Cells.
// Use Cases: Provide a cancel button in a desktop UI that instantly stops massive worksheet‑to‑TIFF export. | Release server resources in a web API when a client aborts a TIFF report request. | Enforce a maximum processing time for TIFF generation and abort if the limit is exceeded.
// AI Prompts: Show how to connect a UI cancel button to the CancellationToken that interrupts Aspose.Cells rendering. | Create a reusable async method that converts a worksheet to TIFF and accepts a CancellationToken for aborting. | Explain the proper pattern for catching CellsException.Interrupted and disposing streams after cancellation.

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to abort a time‑consuming TIFF conversion in Aspose.Cells by linking a CancellationToken to an InterruptMonitor, handling the CellsException Interrupted, and cleaning up resources.
class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook with a large amount of data to make TIFF conversion time‑consuming.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int i = 0; i < 20000; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i}");
            }

            // Create an interrupt monitor and assign it to the workbook.
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            // Set up a cancellation token that the user can trigger.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                // When cancellation is requested, interrupt the current operation.
                cts.Token.Register(() =>
                {
                    Console.WriteLine("Cancellation requested – interrupting TIFF conversion.");
                    monitor.Interrupt();
                });

                // Simulate a user requesting cancellation after 2 seconds.
                Task.Run(() =>
                {
                    Thread.Sleep(2000);
                    cts.Cancel();
                });

                // Configure rendering options for TIFF output.
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Tiff,
                    OnePagePerSheet = false
                };

                // Create the sheet renderer.
                SheetRender renderer = new SheetRender(sheet, renderOptions);

                // Render the worksheet to a TIFF file using a stream.
                using (FileStream tiffStream = new FileStream("output.tiff", FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    Console.WriteLine("Starting TIFF conversion...");
                    renderer.ToImage(0, tiffStream); // Render all pages into a single TIFF.
                    Console.WriteLine("TIFF conversion completed successfully.");
                }
            }
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            // The operation was aborted by the interrupt monitor.
            Console.WriteLine("TIFF conversion was aborted due to cancellation.");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
