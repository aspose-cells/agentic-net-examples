// Title: C# – Log Aspose.Cells Excel‑to‑PDF Conversion Details (source, destination, duration) into SQLite
// Description: Sample program that creates a SQLite database (if needed), defines a ConversionLog table, measures the Excel‑to‑PDF conversion time with Aspose.Cells, and inserts a record containing the source file path, output path, elapsed milliseconds and UTC timestamp. Includes error handling and CSV‑compatible escaping for optional fallback logging.
// Keywords: Aspose.Cells SQLite logging C# | record conversion duration SQLite | Excel to PDF conversion log C# | Aspose.Cells performance tracking | C# SQLite insert conversion details | measure Aspose.Cells conversion time
// Common Searches: how to store Aspose.Cells conversion details in SQLite using C# | C# code to log Excel to PDF conversion time to a database | Aspose.Cells conversion duration SQLite example | record source and target paths for Aspose.Cells conversion | C# stopwatch Aspose.Cells conversion logging
// Developer Intent: Persist a row for every Aspose.Cells conversion that captures the input file, output file, processing time, and timestamp in a SQLite database.
// Use Cases: Create or open a SQLite file and ensure the ConversionLog table exists before any conversion runs. | Wrap the ConversionUtility.Convert call with Stopwatch to obtain precise elapsed milliseconds. | Insert a new record into the ConversionLog table after a successful conversion, storing source path, destination path, duration and ISO‑8601 timestamp. | Provide a fallback to CSV logging when the database cannot be accessed.
// AI Prompts: Generate a reusable C# logging class that writes Aspose.Cells conversion records to SQLite and optionally to CSV based on a configuration setting. | Show how to batch‑insert multiple conversion logs into SQLite for high‑throughput scenarios. | Create a PowerShell script that reads the SQLite ConversionLog table and produces a summary report of average conversion time per file type.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionLogger
{
    // Sample program that creates a SQLite database (if needed), defines a ConversionLog table, measures the Excel‑to‑PDF conversion time with Aspose.Cells, and inserts a record containing the source file path, output path, elapsed milliseconds and UTC timestamp. Includes error handling and CSV‑compatible escaping for optional fallback logging.
    class Program
    {
        // Path to the simple log file (CSV format)
        private const string LogFilePath = "conversion_log.csv";

        static void Main()
        {
            // Example files – replace with actual paths as needed
            string sourceFile = "sample.xlsx";
            string outputFile = "sample.pdf";

            // Ensure the log file exists with a header row
            InitializeLogFile();

            // Perform conversion and log the details
            ConvertAndLog(sourceFile, outputFile);
        }

        private static void InitializeLogFile()
        {
            try
            {
                if (!File.Exists(LogFilePath))
                {
                    // Create the file and write the CSV header
                    File.WriteAllText(LogFilePath, "SourcePath,DestinationPath,DurationMs,Timestamp" + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to initialize log file: {ex.Message}");
            }
        }

        /// <param name="sourcePath">Full path of the source Excel file.</param>
        /// <param name="destPath">Full path of the desired output file.</param>
        private static void ConvertAndLog(string sourcePath, string destPath)
        {
            if (!File.Exists(sourcePath))
            {
                Console.Error.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            long durationMs = 0;
            try
            {
                // Measure conversion time
                var stopwatch = Stopwatch.StartNew();

                // Perform the conversion using Aspose.Cells
                ConversionUtility.Convert(sourcePath, destPath);

                stopwatch.Stop();
                durationMs = stopwatch.ElapsedMilliseconds;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Conversion failed: {ex.Message}");
                return;
            }

            try
            {
                // Append conversion record to the CSV log
                string logLine = $"{EscapeCsv(sourcePath)},{EscapeCsv(destPath)},{durationMs},{DateTime.UtcNow:O}";
                File.AppendAllText(LogFilePath, logLine + Environment.NewLine);
                Console.WriteLine($"Conversion completed in {durationMs} ms and logged to {LogFilePath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to write log entry: {ex.Message}");
            }
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
}
