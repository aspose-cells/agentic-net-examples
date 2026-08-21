// Title: Log Aspose.Cells formula warnings to a file with timestamps using IWarningCallback (C#)
// Description: Shows how to implement a custom IWarningCallback in Aspose.Cells for .NET that writes each warning—type, description, and a timestamp—to a chosen log file. The sample assigns the callback, forces a warning with an unsupported function, runs workbook calculation, and saves both the workbook and the log.
// Keywords: Aspose.Cells IWarningCallback | C# warning log | timestamped warning file | unsupported formula warning | Aspose.Cells workbook calculation | log file Aspose.Cells | Excel warning callback .NET | formula error logging C#
// Common Searches: how to log Aspose.Cells warnings to a text file | IWarningCallback example C# | save formula warnings with timestamps Aspose.Cells | capture unsupported function warnings in .NET | Aspose.Cells warning callback tutorial
// Developer Intent: Create a .NET IWarningCallback that appends each Aspose.Cells warning, including its type and description, to a log file with a date‑time stamp.
// Use Cases: Record warnings from unsupported Excel functions for post‑run analysis. | Maintain an audit trail of workbook calculation issues in automated reporting pipelines. | Provide developers with a searchable log to troubleshoot formula errors during batch processing.
// AI Prompts: Generate a C# class that implements Aspose.Cells IWarningCallback and writes warnings with timestamps to a specified file. | Demonstrate how to attach a custom warning callback to Workbook.Settings.WarningCallback and trigger it using an unsupported formula. | Write a script that reads the generated warnings.log and groups warnings by type after workbook processing.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Custom warning callback that logs warnings to a file with timestamps
    // Shows how to implement a custom IWarningCallback in Aspose.Cells for .NET that writes each warning—type, description, and a timestamp—to a chosen log file. The sample assigns the callback, forces a warning with an unsupported function, runs workbook calculation, and saves both the workbook and the log.
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
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | Type: {warningInfo.Type} | Description: {warningInfo.Description}";
            // Append the entry to the log file
            File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path for the warning log file
            string logPath = "warnings.log";

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set the custom warning callback
            workbook.Settings.WarningCallback = new FileLoggingWarningCallback(logPath);

            // Insert a formula that is likely to generate a warning (unsupported function)
            sheet.Cells["A1"].Formula = "=UNSUPPORTEDFUNC(B1)";

            // Perform a calculation to trigger the warning
            workbook.CalculateFormula();

            // Save the workbook (any additional warnings during save will also be logged)
            workbook.Save("Output.xlsx");

            // Optional: inform the user that processing is complete
            Console.WriteLine("Workbook saved. Warnings (if any) have been logged to " + logPath);
        }
    }
}
