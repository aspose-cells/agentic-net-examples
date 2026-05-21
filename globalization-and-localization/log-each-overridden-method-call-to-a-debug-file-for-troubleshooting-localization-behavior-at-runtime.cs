using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLocalizationLogging
{
    // Custom globalization settings that logs each overridden method call
    public class LoggingGlobalizationSettings : GlobalizationSettings
    {
        private readonly string _logFilePath;

        public LoggingGlobalizationSettings(string logDirectory)
        {
            // Ensure the log directory exists
            Directory.CreateDirectory(logDirectory);
            _logFilePath = Path.Combine(logDirectory, "LocalizationLog.txt");
        }

        // Helper to append a line to the log file with a timestamp
        private void Log(string message)
        {
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}";
            File.AppendAllLines(_logFilePath, new[] { entry });
        }

        // Override to log calls for getting a localized function name
        public override string GetLocalFunctionName(string standardName)
        {
            Log($"GetLocalFunctionName called with standardName=\"{standardName}\"");
            // Use base implementation for actual behavior
            return base.GetLocalFunctionName(standardName);
        }

        // Override to log calls for getting a localized built‑in name
        public override string GetLocalBuiltInName(string standardName)
        {
            Log($"GetLocalBuiltInName called with standardName=\"{standardName}\"");
            return base.GetLocalBuiltInName(standardName);
        }

        // Override to log calls for converting a localized built‑in name back to the standard name
        public override string GetStandardBuiltInName(string localName)
        {
            Log($"GetStandardBuiltInName called with localName=\"{localName}\"");
            return base.GetStandardBuiltInName(localName);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: use provided creation logic)
            Workbook workbook = new Workbook();

            // Set up the logging globalization settings
            string logDir = Path.Combine(Environment.CurrentDirectory, "Logs");
            workbook.Settings.GlobalizationSettings = new LoggingGlobalizationSettings(logDir);

            // Sample data to trigger localization methods
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Use a standard function name – this will invoke GetLocalFunctionName internally
            sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

            // Use a built‑in name that may be localized – this will invoke GetLocalBuiltInName
            // For demonstration, we call the method directly
            var settings = (LoggingGlobalizationSettings)workbook.Settings.GlobalizationSettings;
            string localName = settings.GetLocalBuiltInName("Total");

            // Use the returned local name in a formula to cause GetStandardBuiltInName later
            sheet.Cells["B2"].Formula = $"={localName}(A1:A3)";

            // Calculate formulas to ensure all methods are exercised
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: use provided save logic)
            string outputPath = Path.Combine(Environment.CurrentDirectory, "LocalizedWorkbook.xlsx");
            workbook.Save(outputPath);
        }
    }
}