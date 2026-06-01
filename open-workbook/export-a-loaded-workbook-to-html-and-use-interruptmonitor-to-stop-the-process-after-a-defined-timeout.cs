using System;
using System.IO;
using Aspose.Cells;

class ExportWorkbookToHtmlWithInterrupt
{
    static void Main()
    {
        // Time limits in milliseconds
        const int loadTimeLimit = 2000; // 2 seconds for loading
        const int saveTimeLimit = 1500; // 1.5 seconds for saving

        // Verify that the source workbook exists
        const string sourcePath = "Large.xlsx";
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: Source file \"{sourcePath}\" not found.");
            return;
        }

        // Create an interrupt monitor (false = do not interrupt immediately)
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

        // Configure load options to use the monitor
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        Workbook wb = null;

        try
        {
            // Start monitoring the load operation
            monitor.StartMonitor(loadTimeLimit);

            // Load the workbook with monitoring
            wb = new Workbook(sourcePath, loadOptions);
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Loading was interrupted due to timeout.");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during loading: {ex.Message}");
            return;
        }
        // No explicit StopMonitor call – the monitor stops automatically after the operation

        if (wb == null)
        {
            Console.WriteLine("Workbook could not be loaded.");
            return;
        }

        try
        {
            // Assign the same monitor to the workbook for the save operation
            wb.InterruptMonitor = monitor;

            // Start monitoring the save operation
            monitor.StartMonitor(saveTimeLimit);

            // Export the workbook to HTML
            wb.Save("Result.html", SaveFormat.Html);
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Saving was interrupted due to timeout.");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during saving: {ex.Message}");
            return;
        }
        // No explicit StopMonitor call – the monitor stops automatically after the operation

        Console.WriteLine("Export completed successfully.");
    }
}