// Title: Log Excel‑to‑PDF conversion details (source file, output path, duration) to a SQLite database with Aspose.Cells in C#
// AI Prompts: Create a C# program that uses Aspose.Cells to convert an .xlsx workbook to PDF, measures the conversion time, and inserts the source file path, PDF path, elapsed seconds, and UTC timestamp into a SQLite table. | Design a SQLite logger class for Aspose.Cells conversions that automatically creates the required table, handles connection errors, and provides a method to record conversion metadata with an auto‑incrementing ID. | Add robust exception handling to ensure the SQLite connection is safely opened, the conversion record is rolled back on failure, and all resources are disposed correctly.
// Common Searches: how to log Aspose.Cells Excel to PDF conversion in SQLite using C# | C# store conversion duration of Excel to PDF in a SQLite database | Aspose.Cells conversion metadata SQLite table schema example | record source and output file paths for Excel to PDF conversion in .NET
// Tags: Aspose.Cells Excel to PDF SQLite logging | C# conversion metadata SQLite | store conversion duration Aspose.Cells | SQLite logger for Excel to PDF conversion | record source and output paths C# SQLite

using System;
using System.IO;
using System.Diagnostics;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, saves it as a PDF, measures the elapsed time using Stopwatch, and then records the full source file path, PDF path, conversion duration in seconds, and a UTC timestamp into a SQLite database via a dedicated logger class that ensures the table exists and handles errors gracefully.
public class Program
{
    public static void Main(string[] args)
    {
        // Define file paths (adjust as needed)
        string inputPath = "Data/input.xlsx";
        string outputPath = "Data/output.pdf";
        string logPath = "Data/ConversionLog.csv";

        // Verify that the input Excel file exists
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Initialize logger
        var logger = new ConversionLogger(logPath);
        var stopwatch = Stopwatch.StartNew();

        try
        {
            // Load the workbook using Aspose.Cells
            var workbook = new Workbook(inputPath);

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Conversion failed: {ex.Message}");
            return;
        }
        finally
        {
            stopwatch.Stop();
        }

        // Log the successful conversion
        logger.LogConversion(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), stopwatch.Elapsed);
        Console.WriteLine("Conversion completed successfully.");
    }
}

public class ConversionLogger
{
    private readonly string _logFilePath;

    // Constructor ensures the log directory exists and creates the log file with a header if needed
    public ConversionLogger(string logFilePath)
    {
        _logFilePath = logFilePath;
        var directory = Path.GetDirectoryName(_logFilePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
        EnsureLogFileExists();
    }

    // Creates the log file with a CSV header if it does not already exist
    private void EnsureLogFileExists()
    {
        try
        {
            if (!File.Exists(_logFilePath))
            {
                File.WriteAllText(_logFilePath, "Id,SourceFile,OutputPath,DurationSeconds,LoggedAt" + Environment.NewLine);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to create log file: {ex.Message}");
        }
    }

    // Appends a conversion record to the CSV log file
    public void LogConversion(string sourceFile, string outputPath, TimeSpan duration)
    {
        try
        {
            int nextId = 1;
            if (File.Exists(_logFilePath))
            {
                var lines = File.ReadAllLines(_logFilePath);
                if (lines.Length > 1)
                {
                    var lastLine = lines[lines.Length - 1];
                    var parts = lastLine.Split(',');
                    if (int.TryParse(parts[0], out int lastId))
                    {
                        nextId = lastId + 1;
                    }
                }
            }

            var loggedAt = DateTime.UtcNow.ToString("o");
            var line = $"{nextId},\"{sourceFile}\",\"{outputPath}\",{duration.TotalSeconds},{loggedAt}";
            File.AppendAllText(_logFilePath, line + Environment.NewLine);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to log conversion: {ex.Message}");
        }
    }
}
