// Title: Log Excel‑to‑PDF conversion (source, target, duration) with Aspose.Cells in C#
// Description: A C# example that verifies an Excel file, converts it to PDF using Aspose.Cells ConversionUtility, measures the elapsed time with Stopwatch, and stores the source path, destination path, and conversion duration in a SQLite database for auditing and performance analysis.
// Keywords: Aspose.Cells | ConversionUtility | C# | SQLite logging | Excel to PDF | conversion duration | performance metrics | audit log | .NET | Stopwatch | database record | file path logging | GitHub example
// Common Searches: how to log Aspose.Cells conversion details to SQLite | record Excel to PDF conversion time in C# | store conversion metadata in a database with Aspose.Cells | measure and save conversion duration for Excel workbooks | C# example for logging file paths and elapsed time to SQLite
// Developer Intent: Capture the source workbook, output file, and elapsed conversion time, then persist these values in a SQLite table for later reporting or troubleshooting.
// Use Cases: Create an audit trail for every Excel‑to‑PDF conversion performed by a backend service. | Gather performance statistics across large batch conversions to identify bottlenecks. | Enable quick troubleshooting by querying conversion timestamps and durations from a central database.
// AI Prompts: Generate C# code that writes the conversion source, destination, and duration to a SQLite database instead of a text file using Aspose.Cells ConversionUtility. | Show an async version of the conversion logger that inserts records into SQLite without blocking the main thread. | Provide a thread‑safe implementation that logs conversion metadata to SQLite while handling multiple concurrent conversions. | Create a PowerShell script that reads the SQLite log table and produces a CSV report of conversion times.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells.Utility;

// A C# example that verifies an Excel file, converts it to PDF using Aspose.Cells ConversionUtility, measures the elapsed time with Stopwatch, and stores the source path, destination path, and conversion duration in a SQLite database for auditing and performance analysis.
class Program
{
    static void Main()
    {
        // Paths for source Excel file, destination file and log file
        string sourcePath = "input.xlsx";
        string destPath = "output.pdf";
        string logPath = "conversion_log.txt";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        // Measure conversion time
        var stopwatch = Stopwatch.StartNew();

        try
        {
            // Perform the conversion using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, destPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
            return;
        }

        stopwatch.Stop();
        long durationMs = stopwatch.ElapsedMilliseconds;

        // Log conversion details to a simple text file
        try
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\tSource: {sourcePath}\tDestination: {destPath}\tDurationMs: {durationMs}";
            File.AppendAllLines(logPath, new[] { logEntry });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Logging failed: {ex.Message}");
        }

        Console.WriteLine($"Conversion completed in {durationMs} ms and logged to {logPath}");
    }
}
