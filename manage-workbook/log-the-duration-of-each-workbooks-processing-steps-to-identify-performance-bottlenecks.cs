// Title: Log load, processing, and save durations for each Excel workbook using Aspose.Cells in C#
// Common Searches: how to record load, processing, and save times for Excel files with Aspose.Cells in C# | using Stopwatch to profile workbook operations in Aspose.Cells C# example | log performance metrics of formula calculation and auto‑fit columns in Aspose.Cells | measure individual step durations when processing multiple .xlsx files with Aspose.Cells | create a processing log for batch Excel workbook conversion using Aspose.Cells C#
// Tags: measure workbook load time Aspose.Cells | log Excel processing duration C# | benchmark workbook save performance Aspose.Cells | track formula calculation timing Aspose.Cells | auto‑fit column execution time Aspose.Cells | batch workbook performance logging Aspose.Cells

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The WorkbookProcessor iterates over .xlsx files, uses Stopwatch to capture elapsed milliseconds for loading, recalculating formulas, auto‑fitting columns, and saving each workbook, and writes these timings to a log file for performance analysis.
class WorkbookProcessor
{
    // Path to folder containing workbooks to process
    private readonly string _inputFolder;
    // Path to folder where processed workbooks will be saved
    private readonly string _outputFolder;
    // Path to log file
    private readonly string _logFilePath;

    public WorkbookProcessor(string inputFolder, string outputFolder, string logFilePath)
    {
        _inputFolder = inputFolder;
        _outputFolder = outputFolder;
        _logFilePath = logFilePath;
    }

    public void ProcessAll()
    {
        // Ensure output folder exists
        Directory.CreateDirectory(_outputFolder);
        // Initialize log file
        File.WriteAllText(_logFilePath, $"Workbook processing log - {DateTime.Now}\r\n\r\n");

        // Get all Excel files in the input folder
        var files = Directory.GetFiles(_inputFolder, "*.xlsx");
        foreach (var filePath in files)
        {
            ProcessSingleWorkbook(filePath);
        }
    }

    private void ProcessSingleWorkbook(string inputFilePath)
    {
        var fileName = Path.GetFileName(inputFilePath);
        var outputFilePath = Path.Combine(_outputFolder, fileName);

        // Stopwatch for loading
        var swLoad = Stopwatch.StartNew();
        // --- Load workbook (use provided load rule) ---
        var workbook = new Workbook(inputFilePath);
        swLoad.Stop();

        // Stopwatch for processing (example: calculate formulas, auto-fit columns)
        var swProcess = Stopwatch.StartNew();
        // --- Example processing steps ---
        workbook.CalculateFormula();                     // Recalculate all formulas
        foreach (Worksheet sheet in workbook.Worksheets) // Auto-fit columns for each sheet
        {
            sheet.AutoFitColumns();
        }
        swProcess.Stop();

        // Stopwatch for saving
        var swSave = Stopwatch.StartNew();
        // --- Save workbook (use provided save rule) ---
        workbook.Save(outputFilePath);
        swSave.Stop();

        // Log durations
        var logEntry = $"File: {fileName}\r\n" +
                       $"  Load Time:    {swLoad.ElapsedMilliseconds} ms\r\n" +
                       $"  Process Time: {swProcess.ElapsedMilliseconds} ms\r\n" +
                       $"  Save Time:    {swSave.ElapsedMilliseconds} ms\r\n\r\n";

        File.AppendAllText(_logFilePath, logEntry);
    }
}

// Example usage
class Program
{
    static void Main()
    {
        // Adjust paths as needed
        string inputFolder = @"C:\Workbooks\Input";
        string outputFolder = @"C:\Workbooks\Output";
        string logFile = @"C:\Workbooks\processing_log.txt";

        var processor = new WorkbookProcessor(inputFolder, outputFolder, logFile);
        processor.ProcessAll();

        Console.WriteLine("Processing completed. Check the log file for duration details.");
    }
}
