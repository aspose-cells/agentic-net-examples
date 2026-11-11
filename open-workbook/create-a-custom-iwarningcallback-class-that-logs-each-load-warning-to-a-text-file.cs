// Title: Implement a custom IWarningCallback in Aspose.Cells for C# to log workbook load warnings to a text file
// Common Searches: C# Aspose.Cells how to capture load warnings with a custom IWarningCallback | log Excel workbook load warnings to a file using Aspose.Cells .NET | thread‑safe warning logging implementation for Aspose.Cells workbook loading | Aspose.Cells warning callback example that writes warnings to a text file | save Aspose.Cells load warnings to a custom log file in a C# application
// Tags: Aspose.Cells custom IWarningCallback logging | Excel workbook load warning to text file C# | thread‑safe file append for warning logs | Aspose.Cells warning callback implementation .NET | warning log directory creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Custom warning callback that logs each warning to a text file.
    // The sample defines a FileWarningCallback class that implements IWarningCallback. Its Warning method formats each warning with a timestamp and appends it to a specified log file using a lock for thread safety. In Main, the callback is assigned to Workbook.Settings.WarningCallback before loading the workbook, ensuring all load warnings are recorded to the text file while the workbook is processed and saved.
    public class FileWarningCallback : IWarningCallback
    {
        private readonly string _logFilePath;
        private readonly object _lockObj = new object();

        public FileWarningCallback(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        // This method is called by Aspose.Cells for each warning generated during load.
        public void Warning(WarningInfo info)
        {
            // Build a readable warning message.
            string message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] " +
                             $"Warning Type: {info.Type}, " +
                             $"Message: {info.Description}";

            // Append the warning to the log file in a thread‑safe manner.
            lock (_lockObj)
            {
                File.AppendAllText(_logFilePath, message + Environment.NewLine);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the Excel file to load.
            string inputFile = @"C:\Data\Sample.xlsx";

            // Path to the warning log file.
            string warningLogFile = @"C:\Logs\LoadWarnings.txt";

            // Ensure the warning log file exists.
            try
            {
                string? logDir = Path.GetDirectoryName(warningLogFile);
                if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                if (!File.Exists(warningLogFile))
                {
                    File.Create(warningLogFile).Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to prepare log file: {ex.Message}");
                return;
            }

            // Verify the input file exists before loading.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the workbook using the constructor that accepts a file path.
                Workbook workbook = new Workbook(inputFile);

                // Assign the custom warning callback to capture load warnings.
                workbook.Settings.WarningCallback = new FileWarningCallback(warningLogFile);

                // Perform any additional processing on the workbook here...
                // For demonstration, we will just save it to a new file.
                string outputFile = @"C:\Data\Sample_Output.xlsx";

                // Ensure the output directory exists.
                string? outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputFile);
                Console.WriteLine("Workbook processed and saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}
