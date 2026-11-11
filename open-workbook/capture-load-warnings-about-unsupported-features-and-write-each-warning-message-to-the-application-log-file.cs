// Title: Log workbook loading errors and unsupported feature warnings with Aspose.Cells for .NET
// Common Searches: how to log Aspose.Cells workbook load errors in C# | capture unsupported feature warnings when opening an Excel file with Aspose.Cells | write Aspose.Cells load warnings to a text log file in .NET | C# example for logging workbook load exceptions using Aspose.Cells | handle missing LoadWarnings in Aspose.Cells versions prior to 23.12
// Tags: Aspose.Cells workbook load error logging | C# log Excel load warnings | Aspose.Cells LoadOptions unsupported feature handling | write load exceptions to application log | Aspose.Cells version 23.12 load warnings support

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a log directory, checks the Excel file, loads the workbook with Aspose.Cells, and records any loading exceptions or general errors to a specified log file. It notes that LoadWarnings are unavailable before version 23.12 and logs a success message when the workbook loads without exception.
class Program
{
    static void Main()
    {
        // Paths
        string inputPath = @"C:\Data\Sample.xlsx";
        string logPath   = @"C:\Logs\AppLog.txt";

        try
        {
            // Ensure the log directory exists
            string logDir = Path.GetDirectoryName(logPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);
        }
        catch (Exception dirEx)
        {
            // If we cannot create the log directory, write to console and abort
            Console.WriteLine($"Failed to create log directory: {dirEx.Message}");
            return;
        }

        try
        {
            // Verify the input file exists
            if (!File.Exists(inputPath))
            {
                File.AppendAllText(logPath,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Error: Input file not found: {inputPath}{Environment.NewLine}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load workbook (default LoadOptions)
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                // Log loading exceptions (including possible warnings encapsulated as exceptions)
                File.AppendAllText(logPath,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Load Exception: {loadEx.Message}{Environment.NewLine}");
                return;
            }

            // Aspose.Cells versions prior to 23.12 do not expose LoadWarnings directly.
            // If needed, developers can inspect workbook.LoadOptions or handle specific exceptions.
            // Here we simply log that the workbook was loaded successfully.
            File.AppendAllText(logPath,
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Workbook loaded successfully.{Environment.NewLine}");

            // Further processing of the workbook can be placed here...
        }
        catch (Exception ex)
        {
            // Log any unexpected exceptions
            File.AppendAllText(logPath,
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Exception: {ex.Message}{Environment.NewLine}");
        }
    }
}
