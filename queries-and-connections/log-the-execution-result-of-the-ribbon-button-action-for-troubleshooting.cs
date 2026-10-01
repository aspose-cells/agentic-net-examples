// Title: Record success or failure of an Aspose.Cells custom ribbon button in a timestamped log file using C#
// AI Prompts: Write C# code that appends a timestamped success or error message to a text file when an Aspose.Cells ribbon button is clicked. | Add robust exception handling to the ribbon button handler so any thrown exception is captured and logged with a timestamp. | Refactor the LogResult method to write CSV rows containing timestamp, status, and worksheet name for each ribbon button execution.
// Common Searches: C# how to create a log file for Aspose.Cells ribbon button actions | Aspose.Cells add worksheet and write execution result to a text log | timestamped logging of custom Excel ribbon button using Aspose.Cells .NET | exception handling and logging pattern for Aspose.Cells UI ribbon events | save log of ribbon button click to file in Aspose.Cells C# add-in
// Tags: aspocells ribbon button execution logging | timestamped text log c# aspocells | add worksheet and log result aspocells | exception handling aspocells ribbon handler | c# aspocells file logging for excel automation

using Aspose.Cells;
using System;
using System.IO;

// The example defines a RibbonHandler class that ensures a workbook exists, adds a new worksheet named "LogSheet", saves the workbook, and writes a timestamped success or error entry to a text log file, creating the log directory if needed. It demonstrates logging, exception handling, and basic worksheet manipulation with Aspose.Cells in a C# console scenario.
public class RibbonHandler
{
    // Path to the workbook that the ribbon button will work with
    private const string WorkbookPath = @"C:\Temp\Sample.xlsx";

    // Path to the log file where execution results are recorded
    private const string LogPath = @"C:\Temp\RibbonActionLog.txt";

    // Method to be called when the ribbon button is clicked
    public void OnRibbonButtonClick()
    {
        try
        {
            // Ensure the workbook file exists; create a new one if it doesn't
            Workbook workbook;
            if (File.Exists(WorkbookPath))
            {
                workbook = new Workbook(WorkbookPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
                workbook.Save(WorkbookPath);
            }

            // Example operation: add a new worksheet
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "LogSheet";

            // Save the workbook after modification
            workbook.Save(WorkbookPath);

            // Log successful execution
            LogResult($"Ribbon button executed successfully. Added worksheet '{newSheet.Name}' at index {newSheetIndex}.");
        }
        catch (Exception ex)
        {
            // Log any error that occurs during execution
            LogResult($"Ribbon button execution failed: {ex.Message}");
        }
    }

    // Helper method to append a timestamped entry to the log file
    private void LogResult(string message)
    {
        try
        {
            // Ensure the log directory exists
            string logDir = Path.GetDirectoryName(LogPath);
            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}";
            File.AppendAllText(LogPath, entry);
        }
        catch
        {
            // Suppress any logging exceptions to avoid secondary failures
        }
    }
}

// Entry point for the console application
public class Program
{
    public static void Main()
    {
        RibbonHandler handler = new RibbonHandler();
        handler.OnRibbonButtonClick();
    }
}
