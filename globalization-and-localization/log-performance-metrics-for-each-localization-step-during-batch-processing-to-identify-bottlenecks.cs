// Title: Log performance metrics for each localization step in Aspose.Cells batch processing with C#
// Common Searches: C# Aspose.Cells measure load time for each workbook in a batch | How to record localization step duration when processing Excel files with Aspose.Cells | Save performance metrics to CSV while localizing multiple Excel workbooks in C# | Track per‑file processing times (load, replace, save) using Aspose.Cells API | Log batch Excel localization timings to console and CSV with Aspose.Cells
// Tags: Aspose.Cells workbook load timing | Aspose.Cells localization performance logging | Aspose.Cells workbook save latency | C# CSV metrics logger for Excel processing | batch Excel localization metrics collection

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLocalizationMetrics
{
    // Simple logger that writes metrics to console and optionally to a CSV file
    // The example defines a MetricsLogger that writes timestamps, file names, processing steps, and elapsed milliseconds to both the console and a CSV file. It iterates over a list of Excel workbooks, loads each with Aspose.Cells, applies a simple culture‑code replacement, saves the localized file, and logs the duration of the Load, Localization, and Save steps while handling errors.
    public static class MetricsLogger
    {
        private static readonly string LogFilePath = "LocalizationMetrics.csv";
        private static bool _headerWritten = false;

        public static void Log(string fileName, string step, long milliseconds)
        {
            string line = $"{DateTime.Now:O},{fileName},{step},{milliseconds}";
            Console.WriteLine(line);

            // Append to CSV file
            try
            {
                if (!_headerWritten && !File.Exists(LogFilePath))
                {
                    File.WriteAllText(LogFilePath, "Timestamp,FileName,Step,Milliseconds" + Environment.NewLine);
                    _headerWritten = true;
                }
                File.AppendAllText(LogFilePath, line + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write metrics to file: {ex.Message}");
            }
        }

        public static void LogError(string message)
        {
            Console.WriteLine($"Error: {message}");
        }
    }

    class Program
    {
        // Placeholder for actual localization logic.
        // Replace this with real implementation (e.g., replace strings based on culture).
        private static void LocalizeWorkbook(Workbook workbook, string cultureCode)
        {
            // Example: iterate through all worksheets and replace a placeholder token.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("{Locale}"))
                    {
                        // Simple replacement based on culture code.
                        cell.PutValue(cell.StringValue.Replace("{Locale}", cultureCode));
                    }
                }
            }
        }

        static void Main(string[] args)
        {
            // List of source Excel files to process.
            List<string> sourceFiles = new List<string>
            {
                "Report_EN.xlsx",
                "Report_FR.xlsx",
                "Report_DE.xlsx"
                // Add more file paths as needed.
            };

            // Target culture for localization (could be derived per file).
            string targetCulture = "es-ES";

            // Output directory for localized files.
            string outputDir = "LocalizedOutputs";
            Directory.CreateDirectory(outputDir);

            foreach (string sourcePath in sourceFiles)
            {
                string fileName = Path.GetFileName(sourcePath);
                if (!File.Exists(sourcePath))
                {
                    MetricsLogger.LogError($"Source file not found: {sourcePath}");
                    continue; // Skip to next file
                }

                Stopwatch sw = new Stopwatch();
                Workbook workbook = null;

                // Load step
                try
                {
                    sw.Start();
                    workbook = new Workbook(sourcePath); // Load using Aspose.Cells
                    sw.Stop();
                    MetricsLogger.Log(fileName, "Load", sw.ElapsedMilliseconds);
                }
                catch (Exception ex)
                {
                    sw.Stop();
                    MetricsLogger.LogError($"Failed to load workbook '{fileName}': {ex.Message}");
                    continue; // Skip further processing for this file
                }

                // Localization step
                try
                {
                    sw.Restart();
                    LocalizeWorkbook(workbook, targetCulture);
                    sw.Stop();
                    MetricsLogger.Log(fileName, "Localization", sw.ElapsedMilliseconds);
                }
                catch (Exception ex)
                {
                    sw.Stop();
                    MetricsLogger.LogError($"Localization error for '{fileName}': {ex.Message}");
                    // Continue to attempt saving the partially processed workbook
                }

                // Save step
                string outputPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(fileName) + $"_{targetCulture}.xlsx");
                try
                {
                    sw.Restart();
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    sw.Stop();
                    MetricsLogger.Log(fileName, "Save", sw.ElapsedMilliseconds);
                }
                catch (Exception ex)
                {
                    sw.Stop();
                    MetricsLogger.LogError($"Failed to save workbook '{fileName}' to '{outputPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed. Metrics logged to console and LocalizationMetrics.csv.");
        }
    }
}
