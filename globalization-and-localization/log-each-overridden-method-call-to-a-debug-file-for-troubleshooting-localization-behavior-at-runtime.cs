using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLocalizationLogging
{
    // Simple configuration holder for logging
    internal static class Config
    {
        public static string LogDirectory { get; set; } = Path.Combine(Environment.CurrentDirectory, "Logs");
    }

    // Custom globalization settings that logs each overridden method call
    public class LoggingGlobalizationSettings : GlobalizationSettings
    {
        // Helper method to write log entries
        private void Log(string methodName, string parameter)
        {
            // Ensure the log directory exists
            if (!Directory.Exists(Config.LogDirectory))
                Directory.CreateDirectory(Config.LogDirectory);

            string logFile = Path.Combine(Config.LogDirectory, "LocalizationDebug.log");
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {methodName} called with parameter: {parameter}";
            File.AppendAllText(logFile, entry + Environment.NewLine);
        }

        // Override to log GetLocalFunctionName calls
        public override string GetLocalFunctionName(string standardName)
        {
            Log(nameof(GetLocalFunctionName), standardName);
            return base.GetLocalFunctionName(standardName);
        }

        // Override to log GetLocalBuiltInName calls
        public override string GetLocalBuiltInName(string standardName)
        {
            Log(nameof(GetLocalBuiltInName), standardName);
            return base.GetLocalBuiltInName(standardName);
        }

        // Override to log GetStandardBuiltInName calls
        public override string GetStandardBuiltInName(string localName)
        {
            Log(nameof(GetStandardBuiltInName), localName);
            return base.GetStandardBuiltInName(localName);
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Ensure the logging directory exists
                if (!Directory.Exists(Config.LogDirectory))
                    Directory.CreateDirectory(Config.LogDirectory);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Assign the custom logging globalization settings
                workbook.Settings.GlobalizationSettings = new LoggingGlobalizationSettings();

                // Sample operations that trigger localization methods
                Worksheet sheet = workbook.Worksheets[0];

                // Use a standard function name; the overridden GetLocalFunctionName will be invoked internally
                sheet.Cells["A1"].Formula = "=SUM(B1:B3)";
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);

                // Calculate to ensure formulas are processed
                workbook.CalculateFormula();

                // Use a built‑in name that may be localized
                sheet.Cells["A2"].Formula = "=AVERAGE(B1:B3)";
                workbook.CalculateFormula();

                // Ensure output directory exists before saving
                string outputPath = Path.Combine(Environment.CurrentDirectory, "LocalizationDemo.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath) ?? Environment.CurrentDirectory;
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Save the workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Log unexpected errors to the same log file
                string logFile = Path.Combine(Config.LogDirectory, "LocalizationError.log");
                string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - ERROR: {ex.Message}";
                File.AppendAllText(logFile, entry + Environment.NewLine);
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}