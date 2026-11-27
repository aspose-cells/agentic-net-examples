// Title: C# batch program to delete empty worksheets from Excel workbooks with Aspose.Cells
// Common Searches: aspocells c# script to delete empty worksheets from all Excel files in a directory | how to programmatically clean Excel workbooks by removing blank sheets using Aspose.Cells | batch remove blank worksheets while preserving original file type with Aspose.Cells .NET | detect and delete empty worksheets in .xls .xlsx .xlsm files using C# | automate removal of empty sheets from multiple workbooks using Aspose.Cells API
// Tags: Aspose.Cells remove empty worksheets batch | C# detect empty sheet MaxDataRow MaxDataColumn | delete blank Excel sheets preserving file format | iterate worksheet collection backward removal Aspose.Cells | process .xlsx .xls .xlsm files programmatically Aspose.Cells | automated workbook cleanup empty sheets .NET

using System;
using System.IO;
using Aspose.Cells;

namespace RemoveEmptyWorksheetsBatch
{
    // A C# console application that scans an input folder for Excel‑compatible files, loads each workbook with Aspose.Cells, removes any worksheet whose MaxDataRow and MaxDataColumn are -1 (indicating no data), and saves the cleaned workbook to an output folder using the original file format while logging processing statistics.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: RemoveEmptyWorksheetsBatch <inputDirectory> <outputDirectory>");
                return;
            }

            string inputDir = args[0];
            string outputDir = args[1];

            if (!Directory.Exists(inputDir))
            {
                Console.WriteLine($"Error: Input directory does not exist: {inputDir}");
                return;
            }

            if (!Directory.Exists(outputDir))
            {
                try
                {
                    Directory.CreateDirectory(outputDir);
                    Console.WriteLine($"Created output directory: {outputDir}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: Unable to create output directory: {ex.Message}");
                    return;
                }
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv" };

            string[] files = Directory.GetFiles(inputDir, "*.*", SearchOption.TopDirectoryOnly);
            int processedCount = 0;
            int cleanedCount = 0;
            int errorCount = 0;

            foreach (string filePath in files)
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                {
                    // Skip non‑Excel files
                    continue;
                }

                Console.WriteLine($"Processing file: {Path.GetFileName(filePath)}");
                try
                {
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        bool anyRemoved = RemoveEmptyWorksheets(workbook);
                        string outputFilePath = Path.Combine(outputDir, Path.GetFileName(filePath));

                        // Save using the same format as the source file
                        workbook.Save(outputFilePath);
                        Console.WriteLine($"Saved cleaned workbook to: {outputFilePath}");

                        processedCount++;
                        if (anyRemoved)
                        {
                            cleanedCount++;
                            Console.WriteLine("Empty worksheets were removed.");
                        }
                        else
                        {
                            Console.WriteLine("No empty worksheets found.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }

                Console.WriteLine(new string('-', 60));
            }

            Console.WriteLine("Batch processing completed.");
            Console.WriteLine($"Total files processed: {processedCount}");
            Console.WriteLine($"Workbooks cleaned (empty sheets removed): {cleanedCount}");
            Console.WriteLine($"Files with errors: {errorCount}");
        }

        /// <summary>
        /// Removes all empty worksheets from the given workbook.
        /// A worksheet is considered empty when it contains no data cells (MaxDataRow &lt; 0 and MaxDataColumn &lt; 0).
        /// </summary>
        /// <param name="workbook">The workbook to clean.</param>
        /// <returns>True if at least one worksheet was removed; otherwise false.</returns>
        private static bool RemoveEmptyWorksheets(Workbook workbook)
        {
            bool removedAny = false;
            WorksheetCollection sheets = workbook.Worksheets;

            // Iterate backwards because removing a sheet shifts subsequent indices
            for (int i = sheets.Count - 1; i >= 0; i--)
            {
                Worksheet sheet = sheets[i];
                // MaxDataRow/MaxDataColumn return -1 when the sheet has no data
                if (sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0)
                {
                    Console.WriteLine($"Removing empty worksheet: '{sheet.Name}' (index {i})");
                    sheets.RemoveAt(i);
                    removedAny = true;
                }
            }

            return removedAny;
        }
    }
}
