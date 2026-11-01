// Title: Benchmark fast versus standard XLSX workbook save performance with Aspose.Cells OoxmlSaveOptions in C#
// Common Searches: how to benchmark Aspose.Cells workbook save speed using C# Stopwatch | compare fast save and default save times with OoxmlSaveOptions in .NET | measure elapsed milliseconds for saving an XLSX file with Aspose.Cells | C# code example for profiling Aspose.Cells save operation performance | performance difference between fast and standard save in Aspose.Cells
// Tags: Aspose.Cells OoxmlSaveOptions save performance testing | C# workbook save time measurement with Stopwatch | XLSX fast save vs standard save comparison Aspose | profiling Aspose.Cells save operation .NET | benchmarking Aspose.Cells save speed C#

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Loads an existing XLSX workbook, saves it twice using default OoxmlSaveOptions (first as a fast save, then after reloading as a standard save), measures each save with Stopwatch, outputs the elapsed milliseconds and indicates which method was faster.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // -------------------- Save with default options (fast) --------------------
            var saveOptionsFast = new OoxmlSaveOptions(SaveFormat.Xlsx);
            var sw = Stopwatch.StartNew();
            try
            {
                workbook.Save("output_fast.xlsx", saveOptionsFast);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during fast save: {ex.Message}");
                return;
            }
            sw.Stop();
            long timeFast = sw.ElapsedMilliseconds;
            Console.WriteLine($"Fast save time: {timeFast} ms");

            // Reload to avoid any caching effects from the previous save
            workbook = new Workbook(inputPath);

            // -------------------- Save with default options (standard) --------------------
            var saveOptionsStandard = new OoxmlSaveOptions(SaveFormat.Xlsx);
            sw.Restart();
            try
            {
                workbook.Save("output_standard.xlsx", saveOptionsStandard);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during standard save: {ex.Message}");
                return;
            }
            sw.Stop();
            long timeStandard = sw.ElapsedMilliseconds;
            Console.WriteLine($"Standard save time: {timeStandard} ms");

            // -------------------- Performance Comparison --------------------
            if (timeFast < timeStandard)
                Console.WriteLine("Fast save is quicker than standard save.");
            else if (timeFast > timeStandard)
                Console.WriteLine("Standard save is quicker than fast save.");
            else
                Console.WriteLine("Both saves took the same amount of time.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
