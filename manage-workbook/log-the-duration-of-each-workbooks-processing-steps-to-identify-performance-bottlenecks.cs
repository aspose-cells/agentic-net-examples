using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPerformanceLogging
{
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            List<string> files = new List<string>
            {
                "Input1.xlsx",
                "Input2.xlsx",
                // Add more file paths as needed
            };

            foreach (string filePath in files)
            {
                Console.WriteLine($"--- Processing workbook: {filePath} ---");

                // Verify that the input file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File not found – {filePath}");
                    Console.WriteLine();
                    continue;
                }

                try
                {
                    // Stopwatch to measure total processing time
                    Stopwatch totalStopwatch = Stopwatch.StartNew();

                    // 1. Load workbook
                    Stopwatch loadStopwatch = Stopwatch.StartNew();
                    Workbook wb = new Workbook(filePath);
                    loadStopwatch.Stop();
                    Console.WriteLine($"Load time: {loadStopwatch.ElapsedMilliseconds} ms");

                    // 2. Calculate formulas (if any)
                    Stopwatch calcStopwatch = Stopwatch.StartNew();
                    wb.CalculateFormula();
                    calcStopwatch.Stop();
                    Console.WriteLine($"Formula calculation time: {calcStopwatch.ElapsedMilliseconds} ms");

                    // 3. Prepare output path and ensure directory exists
                    string outputDirectory = Path.GetDirectoryName(filePath) ?? string.Empty;
                    string outputPath = Path.Combine(
                        outputDirectory,
                        Path.GetFileNameWithoutExtension(filePath) + "_Processed.xlsx");

                    if (!Directory.Exists(outputDirectory))
                    {
                        Directory.CreateDirectory(outputDirectory);
                    }

                    // 4. Save workbook
                    Stopwatch saveStopwatch = Stopwatch.StartNew();
                    wb.Save(outputPath);
                    saveStopwatch.Stop();
                    Console.WriteLine($"Save time: {saveStopwatch.ElapsedMilliseconds} ms");

                    // 5. Total time
                    totalStopwatch.Stop();
                    Console.WriteLine($"Total processing time: {totalStopwatch.ElapsedMilliseconds} ms");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred while processing {filePath}: {ex.Message}");
                }

                Console.WriteLine();
            }
        }
    }
}