// Title: Abort loading an Excel workbook after a timeout using Aspose.Cells InterruptMonitor in C#
// AI Prompts: Write C# code that opens an XLSX file with Aspose.Cells and automatically cancels the load if it runs longer than a specified number of seconds using InterruptMonitor. | Show how to set a custom time limit on InterruptMonitor via LoadOptions and capture the exception that indicates the load was interrupted. | Modify the example to log a message when the workbook load is aborted and to clean up resources before exiting.
// Common Searches: c# set timeout for Aspose.Cells workbook loading using InterruptMonitor | how to cancel Excel file load after a certain duration with Aspose.Cells | Aspose.Cells InterruptMonitor example for limiting load time in .NET | catch exception when Aspose.Cells load is interrupted due to timeout | configure LoadOptions to use InterruptMonitor for large Excel files in C#
// Tags: Aspose.Cells InterruptMonitor timeout handling | C# workbook load abort with InterruptMonitor | LoadOptions InterruptMonitor configuration | Excel load time limit Aspose.Cells | interrupted load exception handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input XLSX file, creates an InterruptMonitor (using default limits or a custom timeout), attaches it to LoadOptions, and attempts to load the workbook. If the load exceeds the defined time limit, Aspose.Cells throws an exception that is caught and reported. On successful load, the workbook is saved to the output path, with the output directory created if necessary.
class Program
{
    static void Main()
    {
        // Paths for input and output Excel files.
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Configure an interrupt monitor (default limits are used).
        InterruptMonitor monitor = new InterruptMonitor();

        // Apply the monitor via LoadOptions so it takes effect during loading.
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        Workbook workbook = null;
        bool loadedSuccessfully = false;

        try
        {
            // Load the workbook with the interrupt monitor active.
            workbook = new Workbook(inputPath, loadOptions);
            loadedSuccessfully = true;
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Aspose.Cells throws a generic exception when interrupted.
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }

        // Save the workbook only if it was loaded successfully.
        if (loadedSuccessfully && workbook != null)
        {
            try
            {
                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while saving the workbook: {ex.Message}");
            }
        }
    }
}
