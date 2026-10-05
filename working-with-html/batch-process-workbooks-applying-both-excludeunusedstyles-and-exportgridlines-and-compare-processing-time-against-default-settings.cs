// Title: Batch process multiple Excel workbooks with Aspose.Cells, remove unused styles, and compare save performance against default settings
// AI Prompts: Create a C# console application that scans a folder for .xlsx files, loads each workbook, saves it once with default options, then calls RemoveUnusedStyles and saves it again, logging the elapsed milliseconds for both saves. | Enhance the existing batch‑processing code to write the filename, default save time, and optimized save time to a CSV file for later performance analysis. | Add robust error handling and timing measurement to a routine that processes many workbooks, stores the default and optimized versions in separate directories, and prints a summary table of the timing results.
// Common Searches: how to measure Aspose.Cells workbook save time in C# | remove unused styles from Excel files using Aspose.Cells batch processing | compare performance of default save versus RemoveUnusedStyles in Aspose.Cells | C# code to process all .xlsx files in a directory with Aspose.Cells | Aspose.Cells save performance optimization for multiple workbooks
// Tags: batch removeunusedstyles Aspose.Cells | benchmark workbook save Aspose.Cells | default vs optimized save comparison Aspose.Cells | Aspose.Cells performance optimization Excel | C# folder iteration Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The program iterates over every .xlsx file in an InputWorkbooks folder, loads each workbook with Aspose.Cells, saves it using default settings, then calls RemoveUnusedStyles and saves a second copy. It records the elapsed milliseconds for both saves, outputs the timings to the console, and places the default and optimized files in separate output directories.
class Program
{
    static void Main()
    {
        // Directories for input and output workbooks
        string inputDir = "InputWorkbooks";
        string outputDefaultDir = "OutputDefault";
        string outputOptimizedDir = "OutputOptimized";

        // Ensure output directories exist
        Directory.CreateDirectory(outputDefaultDir);
        Directory.CreateDirectory(outputOptimizedDir);

        // Verify input directory exists
        if (!Directory.Exists(inputDir))
        {
            Console.WriteLine($"Input directory \"{inputDir}\" does not exist.");
            return;
        }

        // Process each .xlsx file in the input directory
        foreach (string filePath in Directory.GetFiles(inputDir, "*.xlsx"))
        {
            // Guard against missing files (should not happen with GetFiles, but added for safety)
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            string fileName = Path.GetFileNameWithoutExtension(filePath);

            try
            {
                // ---------- Default processing (no special settings) ----------
                Stopwatch swDefault = Stopwatch.StartNew();

                // Load workbook (default load options)
                Workbook wbDefault = new Workbook(filePath);

                // Save workbook with default settings
                string defaultOutputPath = Path.Combine(outputDefaultDir, fileName + "_default.xlsx");
                wbDefault.Save(defaultOutputPath);

                swDefault.Stop();

                // ---------- Optimized processing (exclude unused styles) ----------
                Stopwatch swOptimized = Stopwatch.StartNew();

                // Load workbook again for a fresh instance
                Workbook wbOptimized = new Workbook(filePath);

                // Remove unused styles to reduce file size
                wbOptimized.RemoveUnusedStyles();

                // Note: ExportGridLines setting is not available in this version of Aspose.Cells.
                // Grid lines are exported by default; if needed, adjust via SaveOptions.

                // Save optimized workbook
                string optimizedOutputPath = Path.Combine(outputOptimizedDir, fileName + "_optimized.xlsx");
                wbOptimized.Save(optimizedOutputPath);

                swOptimized.Stop();

                // Output processing times for comparison
                Console.WriteLine($"{fileName}: Default = {swDefault.ElapsedMilliseconds} ms, Optimized = {swOptimized.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file \"{filePath}\": {ex.Message}");
            }
        }
    }
}
