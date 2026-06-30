using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example - logs warning messages with timestamps to a file
public class FileLoggingWarningCallback : IWarningCallback
{
    private readonly string _logFilePath;

    public FileLoggingWarningCallback(string logFilePath)
    {
        _logFilePath = logFilePath;
    }

    // IWarningCallback implementation
    public void Warning(WarningInfo warningInfo)
    {
        // Build log entry with timestamp, warning type and description
        string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | Type: {warningInfo.WarningType} | Description: {warningInfo.Description}";

        // Append the entry to the log file
        try
        {
            File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
        }
        catch (Exception ex)
        {
            // If logging fails, write to console as a fallback
            Console.WriteLine($"Failed to write warning to log file: {ex.Message}");
        }
    }
}

class Program
{
    static void Main()
    {
        // Path to the warning log file
        string logPath = "warnings.log";

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Assign the custom warning callback to capture warnings during processing
        workbook.Settings.WarningCallback = new FileLoggingWarningCallback(logPath);

        // Example: add a formula that may be unsupported in certain formats
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=UNSUPPORTEDFUNC(1,2)";

        // Save the workbook (warnings about unsupported formulas will be logged)
        workbook.Save("Output.xlsx");
    }
}