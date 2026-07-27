// Title: C# – Custom IWarningCallback in Aspose.Cells to Log Workbook Load Warnings to a Text File
// Description: Demonstrates how to implement a FileLoggingWarningCallback class that implements Aspose.Cells.IWarningCallback, writes each load‑time warning (timestamp, type, description) to a specified text file, and wires the callback into LoadOptions for automatic logging when opening an Excel workbook.
// Keywords: Aspose.Cells IWarningCallback C# | log workbook load warnings | custom warning callback .NET | LoadOptions warning logging | Excel warning file logger | Aspose.Cells warning callback example
// Common Searches: Aspose.Cells custom warning callback example | how to log load warnings in Aspose.Cells | C# IWarningCallback write to file | record Excel load warnings Aspose.Cells | LoadOptions warning callback C#
// Developer Intent: Create a C# IWarningCallback that records every warning generated while loading a workbook into a text file.
// Use Cases: Audit unsupported features across many Excel files by collecting load warnings in a central log. | Debug workbook import issues during development with timestamped warning entries. | Generate compliance reports that include all warnings encountered during batch processing.
// AI Prompts: Write C# code for an IWarningCallback that appends warning type, description, and timestamp to a log file. | Show how to configure Aspose.Cells LoadOptions with a custom warning callback and ensure the log file is cleared before each run. | Provide a sample that filters out specific warning types before writing them to the log.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Custom warning callback that logs each warning to a text file
    // Demonstrates how to implement a FileLoggingWarningCallback class that implements Aspose.Cells.IWarningCallback, writes each load‑time warning (timestamp, type, description) to a specified text file, and wires the callback into LoadOptions for automatic logging when opening an Excel workbook.
    public class FileLoggingWarningCallback : IWarningCallback
    {
        private readonly string _logFilePath;

        public FileLoggingWarningCallback(string logFilePath)
        {
            _logFilePath = logFilePath;

            // Ensure the log file exists and is empty at start
            File.WriteAllText(_logFilePath, string.Empty);
        }

        // This method is called by Aspose.Cells for every warning encountered
        public void Warning(WarningInfo warningInfo)
        {
            // Build a log entry with timestamp, warning type and description
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | Type: {warningInfo.Type} | Description: {warningInfo.Description}{Environment.NewLine}";

            // Append the entry to the log file
            File.AppendAllText(_logFilePath, logEntry);
        }
    }

    public class LoadWithWarningLoggingDemo
    {
        public static void Run()
        {
            // Path to the log file where warnings will be recorded
            string warningLogPath = "load_warnings.txt";

            // Create an instance of the custom warning callback
            IWarningCallback warningCallback = new FileLoggingWarningCallback(warningLogPath);

            // Initialize LoadOptions with the warning callback
            LoadOptions loadOptions = new LoadOptions
            {
                WarningCallback = warningCallback
            };

            // Load the workbook using the options (any warnings during load will be logged)
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Optional: Access a cell to trigger potential warnings
            Console.WriteLine($"First cell value: {workbook.Worksheets[0].Cells["A1"].StringValue}");

            // Save the workbook (no additional warning handling needed here)
            workbook.Save("output.xlsx");
        }
    }

    class Program
    {
        static void Main()
        {
            LoadWithWarningLoggingDemo.Run();
        }
    }
}
