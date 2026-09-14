// Title: Measure and log execution time of each localization step (number format, date format, text translation) while batch processing Excel workbooks with Aspose.Cells in C#
// AI Prompts: Add Stopwatch timing around each localization method (ApplyNumberFormatLocalization, ApplyDateFormatLocalization, ApplyTextTranslationLocalization) inside the batch loop and write the elapsed milliseconds to a performance log file. | Extend the program to capture peak memory usage for each localization step using GC.GetTotalMemory and include the memory values alongside the timing information in the log. | Create a reusable helper that accepts a Workbook and a list of Action<Workbook> localization actions, measures each action with Stopwatch, logs the step name, duration, and optional memory usage, and returns a dictionary of metrics.
// Common Searches: how to log execution time of individual Aspose.Cells localization steps in a C# batch process | measure performance of number format and date format changes with Aspose.Cells .NET | record per‑step processing time for Excel workbook localization using Stopwatch | batch translate Excel worksheet text and track timing with Aspose.Cells | C# example for profiling Aspose.Cells workbook localization operations
// Tags: Aspose.Cells stopwatch profiling | C# batch Excel localization performance | measure number format conversion Aspose.Cells | date format localization timing | text translation latency Aspose.Cells | performance log generation .NET

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample iterates over a collection of Excel files, loads each workbook with Aspose.Cells, applies three culture‑specific localization steps (German number format, French date format, English‑to‑German text translation), measures the elapsed milliseconds for each step and the total processing time using Stopwatch, writes the timings and any errors to a PerformanceLog.txt file, and saves the localized workbook to a 'Processed' directory.
class Program
{
    static void Main()
    {
        // List of workbooks to process in batch
        var inputFiles = new List<string> { "file1.xlsx", "file2.xlsx" };

        // Directory for processed files
        string outputDir = "Processed";
        Directory.CreateDirectory(outputDir);

        // Performance log file
        string logPath = Path.Combine(outputDir, "PerformanceLog.txt");
        using (var logWriter = new StreamWriter(logPath, false))
        {
            foreach (var inputPath in inputFiles)
            {
                try
                {
                    // Verify that the input file exists
                    if (!File.Exists(inputPath))
                    {
                        logWriter.WriteLine($"{Path.GetFileName(inputPath)} - Error: File not found.");
                        continue;
                    }

                    // Load workbook (load rule)
                    Workbook workbook = new Workbook(inputPath);

                    // Stopwatch for total processing time of the current file
                    Stopwatch totalStopwatch = Stopwatch.StartNew();

                    // ---- Localization Step 1: Number format ----
                    Stopwatch stepStopwatch = Stopwatch.StartNew();
                    ApplyNumberFormatLocalization(workbook);
                    stepStopwatch.Stop();
                    logWriter.WriteLine($"{Path.GetFileName(inputPath)} - NumberFormatLocalization: {stepStopwatch.ElapsedMilliseconds} ms");

                    // ---- Localization Step 2: Date format ----
                    stepStopwatch.Restart();
                    ApplyDateFormatLocalization(workbook);
                    stepStopwatch.Stop();
                    logWriter.WriteLine($"{Path.GetFileName(inputPath)} - DateFormatLocalization: {stepStopwatch.ElapsedMilliseconds} ms");

                    // ---- Localization Step 3: Text translation ----
                    stepStopwatch.Restart();
                    ApplyTextTranslationLocalization(workbook);
                    stepStopwatch.Stop();
                    logWriter.WriteLine($"{Path.GetFileName(inputPath)} - TextTranslationLocalization: {stepStopwatch.ElapsedMilliseconds} ms");

                    // Save workbook (save rule)
                    string outputPath = Path.Combine(outputDir, Path.GetFileName(inputPath));
                    workbook.Save(outputPath);

                    totalStopwatch.Stop();
                    logWriter.WriteLine($"{Path.GetFileName(inputPath)} - TotalProcessingTime: {totalStopwatch.ElapsedMilliseconds} ms");
                }
                catch (Exception ex)
                {
                    // Log any unexpected errors for the current file
                    logWriter.WriteLine($"{Path.GetFileName(inputPath)} - Exception: {ex.Message}");
                }
                finally
                {
                    logWriter.WriteLine(new string('-', 50));
                }
            }
        }
    }

    // Apply culture‑specific number format (e.g., German format)
    static void ApplyNumberFormatLocalization(Workbook workbook)
    {
        const string germanNumberFormat = "#,##0.00";
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                if (cell.Type == CellValueType.IsNumeric)
                {
                    Style style = cell.GetStyle();
                    style.Custom = germanNumberFormat;
                    cell.SetStyle(style);
                }
            }
        }
    }

    // Apply culture‑specific date format (e.g., French format)
    static void ApplyDateFormatLocalization(Workbook workbook)
    {
        const string frenchDateFormat = "dd/MM/yyyy";
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                if (cell.Type == CellValueType.IsDateTime)
                {
                    Style style = cell.GetStyle();
                    style.Custom = frenchDateFormat;
                    cell.SetStyle(style);
                }
            }
        }
    }

    // Replace static English text with localized equivalents
    static void ApplyTextTranslationLocalization(Workbook workbook)
    {
        var translations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Total", "Gesamt" },
            { "Amount", "Betrag" },
            { "Date", "Datum" }
        };

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue;
                    if (translations.TryGetValue(text, out string localized))
                    {
                        cell.PutValue(localized);
                    }
                }
            }
        }
    }
}
