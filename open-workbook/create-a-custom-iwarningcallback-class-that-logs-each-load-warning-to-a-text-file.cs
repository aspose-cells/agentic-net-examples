// Title: Log Aspose.Cells Workbook Load Warnings to a Text File via Custom IWarningCallback (C#)
// Description: Demonstrates how to create a FileLoggingWarningCallback class that implements IWarningCallback, writes each warning's type, description, and timestamp to a text file, attaches it to LoadOptions, loads a workbook, and disposes the callback safely.
// Keywords: Aspose.Cells | IWarningCallback | C# | load warnings | text file logging | Workbook loading | LoadOptions | StreamWriter | dispose pattern | error monitoring
// Common Searches: Aspose.Cells capture load warnings C# | custom IWarningCallback example | write Aspose.Cells warnings to log file | how to dispose IWarningCallback after workbook load | log workbook load warnings with Aspose.Cells
// Developer Intent: Implement a warning callback that records every workbook load warning to a persistent log.
// Use Cases: Maintain an audit trail of data issues when processing large batches of Excel files. | Generate summary reports of warning types across multiple workbooks for quality control. | Ensure file handles are released promptly by applying the dispose pattern to the callback.
// AI Prompts: Create a C# IWarningCallback that logs warnings to a CSV file with columns for timestamp, type, and description. | Show the steps to attach a custom warning callback to LoadOptions, load a workbook, and correctly dispose the callback afterward. | Explain how to filter warnings by type (e.g., FormulaError) inside a custom IWarningCallback implementation.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Custom warning callback that writes each warning to a text file
    // Demonstrates how to create a FileLoggingWarningCallback class that implements IWarningCallback, writes each warning's type, description, and timestamp to a text file, attaches it to LoadOptions, loads a workbook, and disposes the callback safely.
    public class FileLoggingWarningCallback : IWarningCallback, IDisposable
    {
        private readonly StreamWriter _writer;
        private bool _disposed;

        public FileLoggingWarningCallback(string filePath)
        {
            // Open the file for appending warnings
            _writer = new StreamWriter(filePath, append: true);
        }

        // This method is called by Aspose.Cells for every warning encountered
        public void Warning(WarningInfo warningInfo)
        {
            if (_disposed) return;

            // Log warning type and description
            _writer.WriteLine($"{DateTime.Now:u} - Type: {warningInfo.Type}, Description: {warningInfo.Description}");
            _writer.Flush();
        }

        // Dispose pattern to close the StreamWriter
        public void Dispose()
        {
            if (!_disposed)
            {
                _writer?.Dispose();
                _disposed = true;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the workbook to be loaded
            string inputPath = "input.xlsx";

            // Path to the warning log file
            string logPath = "load_warnings.txt";

            // Initialize load options with the custom warning callback
            var loadOptions = new LoadOptions
            {
                WarningCallback = new FileLoggingWarningCallback(logPath)
            };

            // Load the workbook using the options (warnings will be logged)
            var workbook = new Workbook(inputPath, loadOptions);

            // Optionally, perform operations on the workbook here

            // Save the workbook to a new file (no additional warning handling needed)
            workbook.Save("output.xlsx");

            // Dispose the warning callback to release the file handle
            if (loadOptions.WarningCallback is IDisposable disposableCallback)
            {
                disposableCallback.Dispose();
            }
        }
    }
}
