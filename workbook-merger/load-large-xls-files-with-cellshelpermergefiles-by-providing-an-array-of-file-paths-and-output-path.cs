// Title: Merge large XLS files into a single workbook with Aspose.Cells CellsHelper.MergeFiles (C#)
// Description: C# example that takes an array of large .xls paths, validates them, merges the files using CellsHelper.MergeFiles with a temporary cache, and verifies the resulting .xlsx workbook.
// Keywords: Aspose.Cells | CellsHelper.MergeFiles | C# merge large XLS | merge multiple XLS files | temporary cache workbook merging | verify merged workbook | large spreadsheet consolidation
// Common Searches: how to merge big xls files using aspose.cells | cellshelper.mergefiles c# example | merge multiple xls to xlsx with cache file | remove temporary cache after workbook merge | validate merged workbook aspsoe cells
// Developer Intent: Combine several large .xls workbooks into one .xlsx file efficiently with Aspose.Cells.
// Use Cases: Consolidate quarterly Excel reports from legacy .xls files into a master workbook for executive analysis. | Automate nightly batch merging of high‑volume data sheets while controlling memory via a temporary cache. | Create a single data source from multiple departmental spreadsheets before importing into a BI system.
// AI Prompts: Write C# code that merges an array of .xls files into one .xlsx using CellsHelper.MergeFiles, includes missing‑file checks and deletes the cache file afterward. | Explain how to modify the sample to output the merged workbook as CSV or PDF instead of XLSX. | Suggest best practices for placing and sizing the temporary cache file when merging very large Excel workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that takes an array of large .xls paths, validates them, merges the files using CellsHelper.MergeFiles with a temporary cache, and verifies the resulting .xlsx workbook.
    public class MergeLargeXlsFilesDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Paths of the large XLS files to be merged
            string[] filesToMerge = new string[]
            {
                "LargeFile1.xls",
                "LargeFile2.xls",
                "LargeFile3.xls"
            };

            // Verify that all source files exist
            foreach (var file in filesToMerge)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"Source file not found: {file}");
                    return;
                }
            }

            // Temporary cache file required by MergeFiles
            string cacheFile = "MergeCache.tmp";

            // Destination merged file
            string outputFile = "MergedLargeFiles.xlsx";

            try
            {
                // Merge the specified files into a single workbook
                CellsHelper.MergeFiles(filesToMerge, cacheFile, outputFile);
                Console.WriteLine($"Files merged successfully. Output saved to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error merging files: {ex.Message}");
                return;
            }
            finally
            {
                // Clean up the temporary cache file
                if (File.Exists(cacheFile))
                {
                    try
                    {
                        File.Delete(cacheFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to delete cache file: {ex.Message}");
                    }
                }
            }

            // Verify the merged workbook by loading it
            if (File.Exists(outputFile))
            {
                try
                {
                    // Load the merged workbook
                    Workbook mergedWorkbook = new Workbook(outputFile);
                    Console.WriteLine($"Merged workbook contains {mergedWorkbook.Worksheets.Count} worksheet(s).");

                    // Optionally save a copy in a different format
                    mergedWorkbook.Save("VerifiedMerged.xlsx", SaveFormat.Xlsx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading or saving merged workbook: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Merged output file not found: {outputFile}");
            }
        }
    }
}
