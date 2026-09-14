// Title: C# example: Convert an Excel workbook to PDF with Aspose.Cells and abort the operation after 20 seconds using a timeout task
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, saves it as PDF, and stops the conversion when it exceeds a 20‑second limit by running workbook.Save inside a Task and applying a timeout. | Show how to wrap workbook.Save(Pdf) in a background Task and use Task.Wait with a TimeSpan to enforce a maximum conversion duration without blocking the main thread. | Provide a pattern for monitoring a long‑running Aspose.Cells PDF export and logging a timeout message when the operation runs longer than the allowed time.
// Common Searches: how to set a 20‑second timeout for Aspose.Cells PDF export in C# | run Aspose.Cells workbook.Save to PDF in background task with cancellation support | limit Aspose.Cells Excel to PDF conversion time using .NET Task.Wait | example of aborting Aspose.Cells SaveFormat.Pdf after a time limit | C# timeout handling for long‑running Aspose.Cells conversions
// Tags: Aspose.Cells PDF export timeout | C# task wrapper for workbook.Save | Excel to PDF conversion time limit | background conversion with Aspose.Cells | monitoring long‑running Aspose.Cells operation

using Aspose.Cells;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

// Loads "input.xlsx" with Aspose.Cells, starts workbook.Save to "output.pdf" inside a Task, and uses Task.Wait with a 20‑second TimeSpan to abort the conversion if it exceeds the allowed duration, logging an appropriate timeout message.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
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

        // Run the PDF conversion in a separate task so we can enforce a timeout
        var conversionTask = Task.Run(() =>
        {
            try
            {
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PDF conversion error: {ex.Message}");
            }
        });

        // Wait up to 20 seconds for the conversion to finish
        if (!conversionTask.Wait(TimeSpan.FromSeconds(20)))
        {
            Console.WriteLine("PDF conversion aborted: operation exceeded 20 seconds.");
            // Note: Aspose.Cells does not provide a direct way to cancel an ongoing Save operation.
            // The task will continue in the background; you may choose to ignore its result.
        }
    }
}
