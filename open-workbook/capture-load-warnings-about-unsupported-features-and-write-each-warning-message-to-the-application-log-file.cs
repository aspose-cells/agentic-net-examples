using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    // Custom warning callback that writes each warning to a log file.
    public class LogWarningCallback : IWarningCallback
    {
        private readonly string _logPath;

        public LogWarningCallback(string logPath)
        {
            _logPath = logPath;
            // Ensure the directory for the log file exists.
            Directory.CreateDirectory(Path.GetDirectoryName(_logPath) ?? string.Empty);
            // Initialize the log file.
            File.WriteAllText(_logPath, $"Log started at {DateTime.Now}{Environment.NewLine}");
        }

        // Called by Aspose.Cells for every warning encountered.
        public void Warning(WarningInfo warningInfo)
        {
            string message = $"Warning Type: {warningInfo.Type}, Description: {warningInfo.Description}";
            File.AppendAllText(_logPath, $"{DateTime.Now}: {message}{Environment.NewLine}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the workbook that may contain unsupported features.
            string inputPath = "sample_with_unsupported_features.xlsx";

            // Path to the application log file.
            string logPath = "load_warnings.log";

            try
            {
                // Verify that the input workbook exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Create LoadOptions and assign the custom warning callback.
                LoadOptions loadOptions = new LoadOptions
                {
                    WarningCallback = new LogWarningCallback(logPath)
                };

                // Load the workbook using the options; any load warnings will be captured.
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Optional: further processing of the workbook can be done here.

                Console.WriteLine("Workbook loaded. Load warnings have been written to the log file.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}