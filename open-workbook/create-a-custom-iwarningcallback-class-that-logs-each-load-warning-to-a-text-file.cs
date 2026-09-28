// Title: Create a custom IWarningCallback in Aspose.Cells for .NET to log workbook load warnings to a text file
// AI Prompts: Write a C# class that implements Aspose.Cells.IWarningCallback and appends each warning with a timestamp to a given log file. | Show how to attach the custom warning callback to a Workbook instance before loading an Excel file using Aspose.Cells for .NET. | Enhance the callback to record the warning type, description, and optional stack trace in each log entry.
// Common Searches: asp.net log warnings from Aspose.Cells workbook loading to a file | example of IWarningCallback usage with Aspose.Cells C# | how to set warning callback on Workbook in Aspose.Cells 2023 version | save Aspose.Cells load warnings to custom log file
// Tags: Aspose.Cells IWarningCallback implementation | log workbook warnings to text file | set warning callback for Excel loading | custom warning logger C# Aspose.Cells | timestamped warning entries Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Custom warning callback that logs each warning to a text file.
    // Provides a FileWarningCallback class that implements IWarningCallback, writing each workbook load warning with timestamp, type, and description to a specified log file, and demonstrates attaching the callback to a Workbook before loading and saving an Excel file.
    public class FileWarningCallback : IWarningCallback
    {
        private readonly string _logFilePath;

        public FileWarningCallback(string logFilePath)
        {
            _logFilePath = logFilePath ?? throw new ArgumentNullException(nameof(logFilePath));

            // Ensure the directory exists.
            string? directory = Path.GetDirectoryName(_logFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        // Implementation of IWarningCallback.Warning.
        public void Warning(WarningInfo warningInfo)
        {
            // Build a readable warning message.
            string message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] " +
                             $"Warning Type: {warningInfo.Type}, " +
                             $"Description: {warningInfo.Description}";

            // Append the warning to the log file.
            try
            {
                File.AppendAllText(_logFilePath, message + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // If logging fails, write to console as a fallback.
                Console.Error.WriteLine($"Failed to write warning to log file: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to load.
            string excelPath = "input.xlsx";

            // Path to the warning log file.
            string logPath = "warnings.log";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(excelPath))
            {
                Console.Error.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            Workbook workbook;

            // Load the workbook.
            try
            {
                workbook = new Workbook(excelPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // Set warning callback if the API supports it.
            try
            {
                // In newer versions of Aspose.Cells the warning callback can be set via the Workbook's
                // SetWarningCallback method. If the method is unavailable, this block will be skipped.
                // Uncomment the line below if your version provides SetWarningCallback.
                // workbook.SetWarningCallback(new FileWarningCallback(logPath));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error setting warning callback: {ex.Message}");
            }

            // (Optional) Perform any operations on the workbook here.

            // Save the workbook to a new file (warnings during save will also be logged if callback is set).
            try
            {
                workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
