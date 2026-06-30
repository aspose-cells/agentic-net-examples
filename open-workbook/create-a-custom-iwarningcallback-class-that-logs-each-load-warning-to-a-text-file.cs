using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Author: Aspose.Cells .NET example - custom warning logger
    public class FileWarningLogger : IWarningCallback
    {
        private readonly string _logFilePath;

        public FileWarningLogger(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        // This method is called by Aspose.Cells for each warning encountered during load/save.
        public void Warning(WarningInfo warningInfo)
        {
            // Append warning details to the specified log file.
            try
            {
                using (StreamWriter writer = new StreamWriter(_logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now:u} | Type: {warningInfo.WarningType} | Description: {warningInfo.Description}");
                }
            }
            catch (Exception ex)
            {
                // If logging fails, write to console as a fallback.
                Console.WriteLine($"Failed to write warning to log file: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the Excel file to load.
            string inputPath = "input.xlsx";

            // Path where the workbook will be saved after loading.
            string outputPath = "output.xlsx";

            // Path to the warning log file.
            string logPath = "load_warnings.log";

            // Create load options and assign the custom warning callback.
            LoadOptions loadOptions = new LoadOptions
            {
                WarningCallback = new FileWarningLogger(logPath)
            };

            // Load the workbook with the specified options (warnings will be logged).
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any workbook manipulations here.

            // Save the workbook. No warning callback is needed for saving in this example.
            workbook.Save(outputPath);
        }
    }
}