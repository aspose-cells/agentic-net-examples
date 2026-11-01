// Title: Log Excel‑to‑PDF conversion details (source file, output path, duration) to a CSV file using Aspose.Cells in C#
// AI Prompts: Create a C# class that writes conversion metadata—source workbook path, generated PDF path, elapsed seconds, and UTC timestamp—to a CSV file, automatically creating the file and its folder if missing. | Modify the logger to store the same conversion information in a SQLite database using System.Data.SQLite, including table creation and parameterized inserts. | Implement thread‑safe CSV logging in C# so that multiple conversion tasks can append rows concurrently without corrupting the log.
// Common Searches: how to write Aspose.Cells conversion log to CSV in C# | record Excel to PDF conversion time and file paths in a .NET log | C# append rows with UTC timestamp to a CSV file for file conversions | store conversion metadata in SQLite database using Aspose.Cells C# | thread‑safe way to log multiple file conversions to CSV in .NET
// Tags: Aspose.Cells CSV conversion logger | C# conversion metadata SQLite | concurrent CSV writes .NET | log workbook to PDF duration | automatic log directory creation C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, converts it to PDF with Aspose.Cells, measures the conversion time, and records the source path, output path, duration in seconds, and a UTC timestamp to a CSV log file, automatically handling directory creation and I/O errors.
public class ConversionLogger
{
    private readonly string _logFilePath;

    // Constructor creates (or opens) the CSV log file and ensures the header exists.
    public ConversionLogger(string logFilePath)
    {
        _logFilePath = logFilePath ?? throw new ArgumentNullException(nameof(logFilePath));

        // Ensure the directory for the log file exists.
        var directory = Path.GetDirectoryName(_logFilePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        try
        {
            // If the file does not exist, create it with a header row.
            if (!File.Exists(_logFilePath))
            {
                File.WriteAllText(_logFilePath, "Id,SourceFile,OutputPath,DurationSeconds,TimestampUtc" + Environment.NewLine);
            }
        }
        catch (Exception ex)
        {
            // Re‑throw as an IOException with context.
            throw new IOException($"Failed to initialize log file at '{_logFilePath}'.", ex);
        }
    }

    // Logs a conversion entry with source file, output path and duration (in seconds).
    public void LogConversion(string sourceFile, string outputPath, TimeSpan duration)
    {
        if (string.IsNullOrEmpty(sourceFile))
            throw new ArgumentException("Source file path must be provided.", nameof(sourceFile));
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("Output path must be provided.", nameof(outputPath));

        try
        {
            // Determine the next Id by reading the last line (simple approach for demo purposes).
            long nextId = 1;
            if (File.Exists(_logFilePath))
            {
                var lines = File.ReadAllLines(_logFilePath);
                if (lines.Length > 1) // header + at least one entry
                {
                    var lastLine = lines[lines.Length - 1];
                    var parts = lastLine.Split(',');
                    if (long.TryParse(parts[0], out var lastId))
                    {
                        nextId = lastId + 1;
                    }
                }
            }

            var timestamp = DateTime.UtcNow.ToString("o");
            var csvLine = $"{nextId},\"{sourceFile}\",\"{outputPath}\",{duration.TotalSeconds},{timestamp}";
            File.AppendAllText(_logFilePath, csvLine + Environment.NewLine);
        }
        catch (Exception ex)
        {
            // Re‑throw as an IOException with context.
            throw new IOException($"Failed to write log entry to '{_logFilePath}'.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Prepare paths relative to the executable directory.
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string logPath = Path.Combine(baseDir, "Data", "ConversionLog.csv");
            var logger = new ConversionLogger(logPath);

            string inputPath = Path.Combine(baseDir, "Input", "sample.xlsx");
            string outputPath = Path.Combine(baseDir, "Output", "sample.pdf");

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            if (File.Exists(inputPath))
            {
                var start = DateTime.UtcNow;

                // Load workbook and convert to PDF using Aspose.Cells.
                var workbook = new Workbook(inputPath);
                workbook.Save(outputPath, SaveFormat.Pdf);

                var duration = DateTime.UtcNow - start;
                logger.LogConversion(inputPath, outputPath, duration);
                Console.WriteLine($"Conversion completed in {duration.TotalSeconds:F2} seconds.");
            }
            else
            {
                Console.WriteLine($"Input file not found: {inputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
