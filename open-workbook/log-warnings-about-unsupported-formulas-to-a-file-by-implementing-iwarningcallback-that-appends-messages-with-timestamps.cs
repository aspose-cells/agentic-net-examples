// Title: How to log unsupported formula warnings to a file with a custom IWarningCallback in Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells capture formula warnings and write them to a log file | using IWarningCallback to log Excel formula warnings with timestamps in .NET | example of Aspose.Cells warning callback for unsupported formulas | append warning messages to a text file during workbook calculation Aspose.Cells | log Aspose.Cells warning info to a file with custom callback
// Tags: custom IWarningCallback implementation Aspose.Cells | timestamped warning logging C# | unsupported formula warning capture .NET | Workbook.Settings.WarningCallback usage | log formula warnings to file

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningLogger
{
    // Implements IWarningCallback to capture warnings during workbook processing
    // Demonstrates a custom IWarningCallback that records each warning generated while processing a workbook—such as unsupported formulas—into a specified log file, prefixing entries with an ISO‑8601 timestamp. The callback is assigned to Workbook.Settings.WarningCallback before formula calculation and workbook saving.
    public class FormulaWarningLogger : IWarningCallback
    {
        private readonly string _logFilePath;

        public FormulaWarningLogger(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        // Called for each warning encountered
        public void Warning(WarningInfo info)
        {
            try
            {
                // Log the warning information (fallback to ToString if specific properties are unavailable)
                using (var writer = new StreamWriter(_logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now:O}: {info}");
                }
            }
            catch
            {
                // Suppress any logging failures
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string logPath = "warnings.log";

            try
            {
                // Ensure input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Assign the custom warning callback
                workbook.Settings.WarningCallback = new FormulaWarningLogger(logPath);

                // Perform workbook operations (e.g., calculations)
                workbook.CalculateFormula();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);

                Console.WriteLine("Workbook processed and saved successfully.");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
                try
                {
                    using (var writer = new StreamWriter(logPath, true))
                    {
                        writer.WriteLine($"{DateTime.Now:O}: Exception - {ex}");
                    }
                }
                catch
                {
                    // Suppress any logging failures
                }
            }
        }
    }
}
