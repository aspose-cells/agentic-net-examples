// Title: Generate a CSV log of formula calculation times per worksheet using Aspose.Cells for .NET
// Common Searches: how to record formula calculation duration for each cell with Aspose.Cells C# | export Aspose.Cells formula evaluation timings to a CSV file | measure performance of worksheet formulas using Aspose.Cells .NET | log start and end timestamps of cell.Calculate in Aspose.Cells | create a formula evaluation performance report with Aspose.Cells
// Tags: Aspose.Cells formula timing export | cell.Calculate duration logging | worksheet formula performance CSV | Aspose.Cells calculation profiling | log formula evaluation metrics .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, iterates through every worksheet and each cell that contains a formula, captures UTC start and end timestamps around the cell.Calculate call, computes the elapsed time, and stores these details in a list. After processing all formulas, it writes a CSV file (Worksheet,Cell,StartTime,EndTime,Duration(ms)) to the specified path, creating the output directory if necessary and handling missing files or calculation errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "FormulaEvaluationLog.csv";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect log entries
            var logEntries = new List<LogEntry>();

            // Options for cell calculation (required by Aspose.Cells API)
            var calcOptions = new CalculationOptions();

            // Iterate through all worksheets and cells with formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        var entry = new LogEntry
                        {
                            WorksheetName = sheet.Name,
                            CellName = cell.Name,
                            StartTime = DateTime.UtcNow
                        };

                        try
                        {
                            // Calculate the formula for the current cell
                            cell.Calculate(calcOptions);
                        }
                        catch (Exception calcEx)
                        {
                            Console.WriteLine($"Error calculating cell {cell.Name} in sheet {sheet.Name}: {calcEx.Message}");
                        }

                        entry.EndTime = DateTime.UtcNow;
                        entry.Duration = entry.EndTime - entry.StartTime;
                        logEntries.Add(entry);
                    }
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write the detailed log to a CSV file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                writer.WriteLine("Worksheet,Cell,StartTime,EndTime,Duration(ms)");
                foreach (var e in logEntries)
                {
                    writer.WriteLine($"{e.WorksheetName},{e.CellName},{e.StartTime:O},{e.EndTime:O},{e.Duration.TotalMilliseconds}");
                }
            }

            Console.WriteLine($"Formula evaluation log saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple DTO for log information
    private class LogEntry
    {
        public string WorksheetName { get; set; } = string.Empty;
        public string CellName { get; set; } = string.Empty;
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration { get; set; }
    }
}
