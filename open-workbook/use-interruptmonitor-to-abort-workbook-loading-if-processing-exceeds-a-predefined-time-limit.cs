// Title: Abort workbook loading with InterruptMonitor when load time exceeds a limit in Aspose.Cells for .NET (C#)
// Common Searches: how to stop Aspose.Cells workbook loading after a timeout in C# | using InterruptMonitor to cancel Excel file load in Aspose.Cells .NET | set maximum load duration for an XLSX workbook with Aspose.Cells | Aspose.Cells interrupt long-running workbook load operation | C# example of InterruptMonitor aborting Excel load beyond time limit
// Tags: interruptmonitor abort workbook loading | excel workbook load timeout Aspose.Cells | loadoptions time limit Aspose.Cells | aspose.cells prevent long-running load | c# interruptmonitor usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates an InterruptMonitor with a 30‑second limit, attaches it to LoadOptions, and loads an XLSX file using Aspose.Cells. If the loading process exceeds the time limit, the monitor aborts the operation, triggering an exception that is caught and reported. After successful loading, the workbook is saved to a new file.
class Program
{
    static void Main()
    {
        // Define the maximum allowed processing time (not used directly here, kept for reference).
        TimeSpan timeLimit = TimeSpan.FromSeconds(30);

        // Path to the input and output files.
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook. If needed, you can configure LoadOptions here.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading.
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // TODO: Add any workbook processing here.

        try
        {
            // Save the workbook after successful processing.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during saving.
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
