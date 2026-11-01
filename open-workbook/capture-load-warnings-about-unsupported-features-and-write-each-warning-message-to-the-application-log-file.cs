// Title: Log unsupported feature load warnings from an Excel workbook using Aspose.Cells LoadOptions and IWarningCallback in C#
// Common Searches: how to log load warnings from Aspose.Cells workbook in C# | using LoadOptions WarningCallback to capture unsupported feature warnings in .NET | example of IWarningCallback for Excel file loading with Aspose.Cells | recording Aspose.Cells load warnings to a custom log file | C# code to write Aspose.Cells load warnings to application log
// Tags: Aspose.Cells load warning callback | custom IWarningCallback implementation C# | log unsupported feature warnings .NET | LoadOptions warning callback example | append workbook load warnings to file

using Aspose.Cells;
using System;
using System.IO;

// The sample creates an application log file, configures LoadOptions with a custom IWarningCallback (LoadWarningLogger) that writes each warning's description and a timestamp to the log, loads an Excel workbook using these options, and logs any unexpected exceptions, ensuring all load‑time warnings about unsupported features are captured.
class Program
{
    static void Main()
    {
        // Paths for the workbook to load and the log file
        string workbookPath = "input.xlsx";
        string logPath = "app.log";

        // Ensure the log file exists
        if (!File.Exists(logPath))
        {
            using (File.Create(logPath)) { }
        }

        // Verify that the workbook file exists before attempting to load it
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
            return;
        }

        try
        {
            // Configure load options with a warning callback
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                WarningCallback = new LoadWarningLogger(logPath)
            };

            // Load the workbook using the configured options
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Workbook is now loaded; further processing can be done here
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Exception: {ex.Message}{Environment.NewLine}");
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Implements Aspose.Cells.IWarningCallback to handle load warnings
class LoadWarningLogger : IWarningCallback
{
    private readonly string _logFilePath;

    public LoadWarningLogger(string logFilePath)
    {
        _logFilePath = logFilePath;
    }

    // Called by Aspose.Cells for each warning encountered during loading
    public void Warning(WarningInfo info)
    {
        // Use the Description property for the warning message
        string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {info.Description}{Environment.NewLine}";
        File.AppendAllText(_logFilePath, logEntry);
    }
}
