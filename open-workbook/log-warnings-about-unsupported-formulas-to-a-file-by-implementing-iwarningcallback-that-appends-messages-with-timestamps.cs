using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Custom warning callback that logs warnings to a file with timestamps
    public class FileLoggingWarningCallback : IWarningCallback
    {
        private readonly string _logFilePath;

        public FileLoggingWarningCallback(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        // This method is called by Aspose.Cells whenever a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            // Build log entry
            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string warningType = warningInfo.Type.ToString();
            string description = warningInfo.Description;
            string logEntry = $"{timeStamp} | Type: {warningType} | Description: {description}{Environment.NewLine}";

            // Append to the log file
            try
            {
                File.AppendAllText(_logFilePath, logEntry);
            }
            catch (Exception ex)
            {
                // If logging fails, write to console as a fallback
                Console.WriteLine($"Failed to write warning to log file: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path for the warning log file
            string warningLogPath = "warnings.log";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula that is likely to generate a warning (unsupported function)
            sheet.Cells["A1"].Formula = "=UNKNOWNFUNC(B1)";

            // Assign the custom warning callback to the workbook settings
            workbook.Settings.WarningCallback = new FileLoggingWarningCallback(warningLogPath);

            // Save the workbook (any warnings during save will trigger the callback)
            workbook.Save("UnsupportedFormulaDemo.xlsx");

            // Inform the user that processing is complete
            Console.WriteLine("Workbook saved. Check the warning log for any messages.");
        }
    }
}