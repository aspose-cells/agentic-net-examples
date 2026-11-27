// Title: Measure and log load, localization, and save times for Excel workbooks in a batch using Aspose.Cells and C#
// Common Searches: how to benchmark Aspose.Cells workbook loading time in C# | record step‑by‑step localization performance for Excel files with Aspose.Cells | write processing metrics of Excel localization batch to CSV in .NET | measure and log save duration of a localized workbook using Aspose.Cells | C# stopwatch usage for tracking Aspose.Cells batch localization steps | performance logging for multi‑language Excel generation with Aspose.Cells
// Tags: Aspose.Cells workbook load timing | Excel localization performance metrics C# | batch processing stopwatch Aspose.Cells | CSV logging of Excel processing times | stepwise timing Aspose.Cells localization | measure save duration Aspose.Cells

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example iterates over a collection of Excel files, loads each workbook with Aspose.Cells, replaces placeholder strings using a dictionary, saves the localized copy, and uses Stopwatch objects to capture load, localization, save, and total execution times. These timings are appended as rows to a CSV file for later analysis of batch processing bottlenecks.
class LocalizationBatchProcessor
{
    // Path to the CSV file where performance metrics will be stored
    private const string MetricsLogPath = "LocalizationMetrics.csv";

    // Entry point
    static void Main()
    {
        // Example input: list of Excel files to process
        var filesToProcess = new List<string>
        {
            @"C:\Data\Report_EN.xlsx",
            @"C:\Data\Report_FR.xlsx",
            // Add more file paths as needed
        };

        // Example localization dictionary: key = placeholder, value = localized text
        var localizationMap = new Dictionary<string, string>
        {
            { "{CompanyName}", "Acme Corp" },
            { "{Year}", DateTime.Now.Year.ToString() }
            // Add more placeholders as needed
        };

        // Ensure the metrics log file has a header row
        InitializeMetricsLog();

        // Process each file and record metrics
        foreach (var filePath in filesToProcess)
        {
            try
            {
                ProcessFile(filePath, localizationMap);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed. Metrics saved to " + MetricsLogPath);
    }

    // Creates the CSV header if the log file does not exist
    private static void InitializeMetricsLog()
    {
        if (!File.Exists(MetricsLogPath))
        {
            var header = "FilePath,LoadTimeMs,LocalizationTimeMs,SaveTimeMs,TotalTimeMs";
            File.WriteAllText(MetricsLogPath, header + Environment.NewLine);
        }
    }

    // Processes a single Excel file: load, apply localization, save, and log timings
    private static void ProcessFile(string filePath, Dictionary<string, string> localizationMap)
    {
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        var totalStopwatch = Stopwatch.StartNew();

        // ---------- Load ----------
        var loadStopwatch = Stopwatch.StartNew();
        Workbook workbook;
        try
        {
            workbook = new Workbook(filePath); // Aspose.Cells load
        }
        catch (Exception loadEx)
        {
            Console.WriteLine($"Failed to load workbook '{filePath}': {loadEx.Message}");
            return;
        }
        loadStopwatch.Stop();

        // ---------- Localization ----------
        var localizationStopwatch = Stopwatch.StartNew();
        ApplyLocalization(workbook, localizationMap);
        localizationStopwatch.Stop();

        // ---------- Save ----------
        var saveStopwatch = Stopwatch.StartNew();
        var outputPath = GetLocalizedFilePath(filePath);
        try
        {
            workbook.Save(outputPath); // Aspose.Cells save
        }
        catch (Exception saveEx)
        {
            Console.WriteLine($"Failed to save workbook '{outputPath}': {saveEx.Message}");
            return;
        }
        saveStopwatch.Stop();

        totalStopwatch.Stop();

        // Log metrics to CSV
        LogMetrics(
            filePath,
            loadStopwatch.ElapsedMilliseconds,
            localizationStopwatch.ElapsedMilliseconds,
            saveStopwatch.ElapsedMilliseconds,
            totalStopwatch.ElapsedMilliseconds);
    }

    // Replaces placeholders in all worksheets with localized values
    private static void ApplyLocalization(Workbook workbook, Dictionary<string, string> localizationMap)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString && cell.StringValue != null)
                    {
                        string updatedValue = cell.StringValue;
                        foreach (var kvp in localizationMap)
                        {
                            if (updatedValue.Contains(kvp.Key))
                            {
                                updatedValue = updatedValue.Replace(kvp.Key, kvp.Value);
                            }
                        }

                        if (updatedValue != cell.StringValue)
                        {
                            cell.PutValue(updatedValue);
                        }
                    }
                }
            }
        }
    }

    // Generates a new file name for the localized workbook
    private static string GetLocalizedFilePath(string originalPath)
    {
        var directory = Path.GetDirectoryName(originalPath);
        var fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
        var extension = Path.GetExtension(originalPath);
        var localizedFileName = $"{fileNameWithoutExt}_Localized{extension}";
        return Path.Combine(directory ?? string.Empty, localizedFileName);
    }

    // Appends a line with the performance data to the CSV log
    private static void LogMetrics(string filePath, long loadMs, long localizationMs, long saveMs, long totalMs)
    {
        var line = $"{EscapeCsv(filePath)},{loadMs},{localizationMs},{saveMs},{totalMs}";
        File.AppendAllText(MetricsLogPath, line + Environment.NewLine);
    }

    // Simple CSV escaping for commas and quotes
    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
