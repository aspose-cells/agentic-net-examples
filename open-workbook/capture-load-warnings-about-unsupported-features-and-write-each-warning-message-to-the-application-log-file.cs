using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example

class Program
{
    static void Main()
    {
        // Path to the log file where warnings will be recorded
        const string logPath = "warnings.log";

        // Clear any existing log content
        File.WriteAllText(logPath, string.Empty);

        // Create a custom warning callback that writes warnings to the log file
        IWarningCallback warningCallback = new FileWarningCallback(logPath);

        // Initialize load options and assign the warning callback
        LoadOptions loadOptions = new LoadOptions
        {
            WarningCallback = warningCallback
        };

        // Load the workbook with the specified options
        // Replace "example.xlsx" with the actual file you want to load
        Workbook workbook = new Workbook("example.xlsx", loadOptions);

        // Workbook is now loaded; any unsupported‑feature warnings have been logged
        Console.WriteLine("Workbook loaded. Check the warnings.log file for details.");
    }
}

// Custom implementation of IWarningCallback that logs each warning message to a file
class FileWarningCallback : IWarningCallback
{
    private readonly string _logFilePath;

    public FileWarningCallback(string logFilePath)
    {
        _logFilePath = logFilePath;
    }

    public void Warning(WarningInfo warningInfo)
    {
        // Compose a readable warning message
        string message = $"Warning: {warningInfo.WarningType} - {warningInfo.Description}";

        // Append the message to the log file
        File.AppendAllText(_logFilePath, message + Environment.NewLine);
    }
}