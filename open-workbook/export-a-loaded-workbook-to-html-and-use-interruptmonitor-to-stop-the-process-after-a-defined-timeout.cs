// Title: Export Excel to HTML with Timeout Using Aspose.Cells InterruptMonitor (C#)
// Description: Demonstrates how to load a large .xlsx file and export it to HTML in .NET while enforcing time limits. A SystemTimeInterruptMonitor is attached to LoadOptions for a 2‑second load timeout and later to the Workbook for a 1.5‑second save timeout. The code catches CellsException.Interrupted to handle aborted operations and safely stops the monitor when possible.
// Keywords: Aspose.Cells | InterruptMonitor | SystemTimeInterruptMonitor | C# HTML export | Excel timeout | LoadOptions | SaveFormat.Html | .NET | CellsException.Interrupted
// Common Searches: Aspose.Cells limit workbook load time | How to abort Excel to HTML conversion after timeout | C# interrupt monitor for large Excel files | Set timeout for Aspose.Cells Save operation | Cancel Aspose.Cells processing with InterruptMonitor
// Developer Intent: Load a workbook and export it to HTML, aborting the operation if loading or saving exceeds predefined time limits.
// Use Cases: Prevent long startup delays by stopping the load of a massive .xlsx file after 2 seconds. | Keep web response times low by terminating the HTML export if it runs longer than 1.5 seconds. | Gracefully handle time‑out interruptions by catching CellsException.Interrupted and providing user feedback.
// AI Prompts: Generate C# code that uses Aspose.Cells SystemTimeInterruptMonitor to limit workbook loading to 3 seconds and saving to PDF to 2 seconds. | Explain how to implement a fallback when InterruptMonitor.StopMonitor is unavailable in older Aspose.Cells versions. | Create a unit test that verifies the interrupt monitor throws an exception when the load operation exceeds the timeout.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load a large .xlsx file and export it to HTML in .NET while enforcing time limits. A SystemTimeInterruptMonitor is attached to LoadOptions for a 2‑second load timeout and later to the Workbook for a 1.5‑second save timeout. The code catches CellsException.Interrupted to handle aborted operations and safely stops the monitor when possible.
class ExportWorkbookToHtmlWithInterrupt
{
    static void Main()
    {
        // Time limits in milliseconds
        const int loadTimeLimit = 2000;   // 2 seconds for loading
        const int saveTimeLimit = 1500;   // 1.5 seconds for saving

        // Verify that the source workbook exists
        const string sourcePath = "Large.xlsx";
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: Input file \"{sourcePath}\" not found.");
            return;
        }

        // Create an interrupt monitor (throws exception when interrupted)
        var monitor = new SystemTimeInterruptMonitor(false);

        // Prepare load options with the monitor
        var loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        Workbook workbook = null;

        try
        {
            // Start monitoring the load operation
            monitor.StartMonitor(loadTimeLimit);

            // Load the workbook
            workbook = new Workbook(sourcePath, loadOptions);
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Loading was interrupted due to timeout.");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during loading: " + ex.Message);
            return;
        }

        // Assign the same monitor to the workbook for the save operation
        workbook.InterruptMonitor = monitor;

        try
        {
            // Start monitoring the save (HTML export) operation
            monitor.StartMonitor(saveTimeLimit);

            // Export the workbook to HTML
            const string resultPath = "Result.html";
            workbook.Save(resultPath, SaveFormat.Html);
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Saving was interrupted due to timeout.");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during saving: " + ex.Message);
            return;
        }
        finally
        {
            // Ensure the monitor is stopped (if supported) without throwing if method is absent
            try
            {
                // Some versions of Aspose.Cells may not expose StopMonitor; ignore if unavailable
                var stopMethod = monitor.GetType().GetMethod("StopMonitor");
                stopMethod?.Invoke(monitor, null);
            }
            catch
            {
                // Ignored – monitor may already be stopped or method not present
            }
        }

        Console.WriteLine("Export completed successfully.");
    }
}
