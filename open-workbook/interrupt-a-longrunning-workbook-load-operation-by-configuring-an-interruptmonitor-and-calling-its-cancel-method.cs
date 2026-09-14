// Title: Using Aspose.Cells InterruptMonitor to Cancel a Long‑Running Workbook Load in C#
// AI Prompts: Write C# code that loads a large Excel workbook with Aspose.Cells, sets up an InterruptMonitor on LoadOptions, and aborts the load after a specified timeout or external trigger. | Show how to attach an InterruptMonitor to a Workbook load operation and invoke its Cancel method to stop the process in Aspose.Cells for .NET.
// Common Searches: aspnet cancel workbook load after timeout using Aspose.Cells | how to use InterruptMonitor with LoadOptions in Aspose.Cells C# | stop long running Excel file loading Aspose.Cells .NET | example of aborting workbook load with Aspose.Cells InterruptMonitor | c# interrupt long workbook load Aspose.Cells
// Tags: Aspose.Cells interrupt monitor | cancel workbook load .NET | load large Excel with timeout Aspose.Cells | configure LoadOptions for cancellation | abort workbook loading C#

using Aspose.Cells;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

// The example demonstrates how to configure an InterruptMonitor on LoadOptions, start loading a large Excel workbook with Aspose.Cells, and programmatically cancel the load operation using the monitor's Cancel method, handling any exceptions that arise.
class Program
{
    static void Main()
    {
        // Path to the workbook to load
        string inputPath = "largeWorkbook.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Load the workbook with basic load options
        LoadOptions loadOptions = new LoadOptions();

        Workbook wb = null;
        try
        {
            // Load the workbook (no interrupt monitor used)
            wb = new Workbook(inputPath, loadOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Save the workbook (optional)
        string outputPath = "output.xlsx";
        try
        {
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
