using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineBenchmark
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source data for the sparklines (same range for all sparklines)
                // Fill cells A1:D1 with sample values
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(1);
                sheet.Cells["D1"].PutValue(3);

                // Define the location of the first sparkline (single cell)
                // Subsequent sparklines will be added row‑by‑row
                CellArea firstLocation = new CellArea
                {
                    StartRow = 0,   // Row 1 (zero‑based)
                    EndRow = 0,
                    StartColumn = 5, // Column F (zero‑based)
                    EndColumn = 5
                };

                // Add a sparkline group with a valid single‑cell location
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, firstLocation);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Benchmark the creation of 5,000 sparklines
                Stopwatch sw = Stopwatch.StartNew();

                for (int i = 0; i < 5000; i++)
                {
                    // Add a sparkline for each row; data range is the same for all
                    // Location varies by row (column F, rows 1‑5000)
                    group.Sparklines.Add("A1:D1", i, 5);
                }

                sw.Stop();
                Console.WriteLine($"Time to create 5,000 sparklines: {sw.ElapsedMilliseconds} ms");

                // Save the workbook (ensure the directory exists)
                string outputPath = "SparklineBenchmark.xlsx";
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}