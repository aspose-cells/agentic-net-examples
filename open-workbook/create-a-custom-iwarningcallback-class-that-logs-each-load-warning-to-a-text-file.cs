using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Custom warning callback that appends each warning to a text file
    public class FileLoggingWarningCallback : IWarningCallback
    {
        private readonly string _logFilePath;

        public FileLoggingWarningCallback(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        // This method is called by Aspose.Cells for every warning encountered during loading
        public void Warning(WarningInfo warningInfo)
        {
            // Build a log entry containing timestamp, warning type and description
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {warningInfo.Type} | {warningInfo.Description}";
            // Append the entry to the log file
            File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
        }
    }

    class Program
    {
        static void Main()
        {
            // Paths for the source workbook and the warning log file
            string sourceFile = "input.xlsx";
            string warningLogFile = "load_warnings.txt";

            // Ensure the log file starts empty
            File.WriteAllText(warningLogFile, string.Empty);

            // Create LoadOptions and assign the custom warning callback
            LoadOptions loadOptions = new LoadOptions
            {
                WarningCallback = new FileLoggingWarningCallback(warningLogFile)
            };

            // Load the workbook using the options that contain the warning callback
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // (Optional) Perform any operations on the workbook here

            // Save the workbook to a new file (no additional warnings are expected here)
            workbook.Save("output.xlsx");
        }
    }
}