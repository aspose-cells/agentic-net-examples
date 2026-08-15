// Title: C# – Log GlobalizationSettings Overrides in Aspose.Cells to Debug Localization at Runtime
// Description: Shows how to derive from Aspose.Cells.GlobalizationSettings, override GetLocalFunctionName, GetLocalBuiltInName, and GetStandardBuiltInName, and write each call with its result to a timestamped log file. The example creates a workbook, assigns the custom settings, adds values, sets a SUM formula to trigger the overrides, runs formula calculation, captures errors, and saves the workbook.
// Keywords: Aspose.Cells | .NET | GlobalizationSettings | logging | localization debugging | override GetLocalFunctionName | override GetLocalBuiltInName | override GetStandardBuiltInName | formula localization | C# example | debug workbook
// Common Searches: Aspose.Cells log GlobalizationSettings overrides | C# debug localization in Aspose.Cells | How to capture GetLocalFunctionName calls | Custom GlobalizationSettings example for Aspose.Cells | Trace formula localization errors .NET
// Developer Intent: Record every invocation of overridden globalization methods in a debug file while processing workbooks with Aspose.Cells.
// Use Cases: Diagnose mismatches between localized and standard function names during formula evaluation. | Identify which built‑in functions are being translated in a multilingual workbook. | Capture localization‑related errors for audit or support purposes.
// AI Prompts: Extend LoggingGlobalizationSettings to also log GetLocalNumberFormat calls in C#. | Configure Aspose.Cells to rotate the localization_debug.log daily and store logs in a user‑specified folder. | Provide a script that parses localization_debug.log and summarizes function name translation failures.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLocalizationLogging
{
    // Custom globalization settings that logs each overridden method call
    // Shows how to derive from Aspose.Cells.GlobalizationSettings, override GetLocalFunctionName, GetLocalBuiltInName, and GetStandardBuiltInName, and write each call with its result to a timestamped log file. The example creates a workbook, assigns the custom settings, adds values, sets a SUM formula to trigger the overrides, runs formula calculation, captures errors, and saves the workbook.
    public class LoggingGlobalizationSettings : GlobalizationSettings
    {
        private readonly string _logFilePath;

        public LoggingGlobalizationSettings(string logDirectory)
        {
            // Ensure the log directory exists
            Directory.CreateDirectory(logDirectory);
            _logFilePath = Path.Combine(logDirectory, "localization_debug.log");
        }

        private void Log(string message)
        {
            // Append the log message with timestamp
            File.AppendAllText(_logFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}{Environment.NewLine}");
        }

        public override string GetLocalFunctionName(string standardName)
        {
            Log($"GetLocalFunctionName called with standardName = '{standardName}'");
            string result = base.GetLocalFunctionName(standardName);
            Log($"GetLocalFunctionName returned '{result}'");
            return result;
        }

        public override string GetLocalBuiltInName(string standardName)
        {
            Log($"GetLocalBuiltInName called with standardName = '{standardName}'");
            string result = base.GetLocalBuiltInName(standardName);
            Log($"GetLocalBuiltInName returned '{result}'");
            return result;
        }

        public override string GetStandardBuiltInName(string localName)
        {
            Log($"GetStandardBuiltInName called with localName = '{localName}'");
            string result = base.GetStandardBuiltInName(localName);
            Log($"GetStandardBuiltInName returned '{result}'");
            return result;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Define a directory for logging
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AsposeLogs");
                var loggingSettings = new LoggingGlobalizationSettings(logDir);

                // Create a new workbook and assign the custom globalization settings
                var workbook = new Workbook();
                workbook.Settings.GlobalizationSettings = loggingSettings;

                // Populate cells
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);

                // Use a formula that will trigger GetLocalFunctionName.
                // Using the standard SUM ensures the formula is valid; the overridden method will still be logged.
                sheet.Cells["A1"].Formula = "=SUM(B1:B2)";

                // Calculate formulas (any errors are caught below)
                try
                {
                    workbook.CalculateFormula();
                }
                catch (Exception ex)
                {
                    // Log calculation errors but continue execution
                    File.AppendAllText(Path.Combine(logDir, "calculation_errors.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - Formula calculation error: {ex.Message}{Environment.NewLine}");
                }

                // Save the workbook
                string outputPath = "LocalizationLoggingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occurred: {e.Message}");
            }
        }
    }
}
