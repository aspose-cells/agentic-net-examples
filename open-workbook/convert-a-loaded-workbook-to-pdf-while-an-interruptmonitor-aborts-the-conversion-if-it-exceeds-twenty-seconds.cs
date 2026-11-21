// Title: Convert an Excel workbook to PDF with a 20‑second timeout using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert Excel to PDF with cancellation after 20 seconds | Set a timeout for workbook.Save to PDF in Aspose.Cells .NET | Abort Aspose.Cells PDF conversion using CancellationToken in C# | How to limit Aspose.Cells PDF export duration to 20 seconds | Implement time‑bound Excel to PDF conversion with Aspose.Cells in .NET
// Tags: Aspose.Cells PDF export timeout | C# workbook.Save with CancellationToken | Aspose.Cells PdfSaveOptions time limit | Excel to PDF conversion abort .NET | Task.Wait timeout for Aspose.Cells conversion

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures PdfSaveOptions, and runs workbook.Save inside a Task. It waits up to 20 seconds for the conversion; if the operation exceeds this limit, the CancellationTokenSource is cancelled, aborting the conversion and reporting a timeout. Exceptions are captured and displayed, and resources are cleaned up after execution.
class WorkbookToPdfWithTimeout
{
    static void Main()
    {
        // Paths for input workbook and output PDF
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use a task with timeout to limit conversion time (20 seconds)
        CancellationTokenSource cts = new CancellationTokenSource();
        Task conversionTask = Task.Run(() =>
        {
            try
            {
                workbook.Save(outputPath, pdfOptions);
            }
            catch (Exception ex)
            {
                // Propagate exception to the outer task
                throw new InvalidOperationException("Conversion failed.", ex);
            }
        }, cts.Token);

        try
        {
            // Wait for the task to complete within the timeout period
            bool completed = conversionTask.Wait(TimeSpan.FromSeconds(20));
            if (completed)
            {
                Console.WriteLine("Workbook successfully converted to PDF.");
            }
            else
            {
                // Timeout occurred; cancel the task if possible
                cts.Cancel();
                Console.WriteLine("Conversion timed out after 20 seconds.");
            }
        }
        catch (AggregateException aggEx)
        {
            // Handle exceptions thrown during conversion
            foreach (var ex in aggEx.InnerExceptions)
            {
                Console.WriteLine($"Conversion error: {ex.Message}");
            }
        }
        finally
        {
            // Clean up resources
            cts.Dispose();
        }
    }
}
