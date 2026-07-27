// Title: Log Aspose.Cells Formula Warnings to a File with Timestamps (C# IWarningCallback)
// Description: Demonstrates a custom FileLoggingWarningCallback that implements IWarningCallback. Each warning is formatted with a date‑time stamp, type, and description, then appended to a specified log file. The callback is assigned to Workbook.Settings.WarningCallback, and an unsupported formula triggers a logged warning when the workbook is saved.
// Keywords: Aspose.Cells | IWarningCallback | C# warning logger | log workbook warnings | timestamped log | unsupported formula | custom warning callback | Aspose.Cells logging example | file append warnings
// Common Searches: Aspose.Cells log warnings to file | C# IWarningCallback example | capture unsupported formula warnings Aspose | timestamped warning log C# | save workbook with warning callback
// Developer Intent: Create a custom IWarningCallback that writes each warning, with a timestamp, to a log file during workbook operations.
// Use Cases: Record formula warnings for audit compliance | Debug spreadsheets by reviewing chronological warning logs | Integrate warning logging into automated Excel processing pipelines | Monitor unsupported functions during large batch conversions
// AI Prompts: Write C# code that implements IWarningCallback to append warning type, description, and ISO‑8601 timestamp to a log file using Aspose.Cells. | Show how to set Workbook.Settings.WarningCallback to a custom file logger and trigger a warning with an unknown function. | Explain how to extend the logger to rotate daily log files while preserving existing entries. | Provide a PowerShell script that runs the compiled example and displays the generated warnings.log content.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Custom warning callback that logs warnings to a file with timestamps
    // Demonstrates a custom FileLoggingWarningCallback that implements IWarningCallback. Each warning is formatted with a date‑time stamp, type, and description, then appended to a specified log file. The callback is assigned to Workbook.Settings.WarningCallback, and an unsupported formula triggers a logged warning when the workbook is saved.
    public class FileLoggingWarningCallback : IWarningCallback
    {
        private readonly string _logFilePath;

        public FileLoggingWarningCallback(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        public void Warning(WarningInfo warningInfo)
        {
            // Build log entry with timestamp, warning type and description
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Type: {warningInfo.Type} - {warningInfo.Description}{Environment.NewLine}";

            // Append the log entry to the specified file
            File.AppendAllText(_logFilePath, logEntry);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path for the warning log file
            string logPath = "warnings.log";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Assign the custom warning callback to the workbook settings
            workbook.Settings.WarningCallback = new FileLoggingWarningCallback(logPath);

            // Example of a formula that may generate a warning (unsupported function)
            worksheet.Cells["A1"].Formula = "=UNKNOWNFUNC(1)";

            // Save the workbook (any warnings during save will be captured by the callback)
            workbook.Save("output.xlsx");
        }
    }
}
