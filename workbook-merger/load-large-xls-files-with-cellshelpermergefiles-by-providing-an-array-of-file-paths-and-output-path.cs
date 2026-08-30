// Title: Combine multiple large XLS workbooks into one file using Aspose.Cells CellsHelper.MergeFiles in C#
// AI Prompts: Generate C# code that merges an array of .xls file paths into a single workbook using Aspose.Cells CellsHelper.MergeFiles with a temporary cache file. | Add logic to verify each source .xls file exists before calling MergeFiles and handle missing files gracefully. | Include post‑merge steps to delete the temporary cache file and load the merged workbook to print the total worksheet count.
// Common Searches: c# how to merge several large .xls workbooks with Aspose.Cells and a cache file | using CellsHelper.MergeFiles to combine multiple Excel files in .NET | validate existence of source XLS files before merging with Aspose.Cells | remove temporary cache after merging Excel workbooks with Aspose.Cells | check worksheet count after merging XLS files using Aspose.Cells
// Tags: Aspose.Cells merge large xls workbooks | CellsHelper.MergeFiles with temporary cache | validate source xls files before merge | cleanup cache file after Excel merge | verify merged workbook worksheet count

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    // The example checks that each large .xls source file exists, then calls CellsHelper.MergeFiles with a temporary cache file to produce a single merged workbook (MergedResult.xls). After merging it deletes the cache file and loads the result to display the number of worksheets.
    public class MergeLargeXlsFiles
    {
        public static void Run()
        {
            // Paths of the large XLS files to be merged
            string[] filesToMerge = new string[]
            {
                "LargeFile1.xls",
                "LargeFile2.xls",
                "LargeFile3.xls"
            };

            // Verify that all source files exist before attempting merge
            foreach (string file in filesToMerge)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"Source file not found: {file}");
                    return;
                }
            }

            // Temporary cache file required by the MergeFiles method
            string cachedFile = "MergeCache.tmp";

            // Destination file that will contain the merged result
            string outputFile = "MergedResult.xls";

            try
            {
                // Merge the specified XLS files into a single workbook
                CellsHelper.MergeFiles(filesToMerge, cachedFile, outputFile);
                Console.WriteLine($"Files merged successfully. Output saved to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error merging files: {ex.Message}");
                return;
            }
            finally
            {
                // Clean up the temporary cache file if it exists
                if (File.Exists(cachedFile))
                {
                    try
                    {
                        File.Delete(cachedFile);
                    }
                    catch (Exception cleanupEx)
                    {
                        Console.WriteLine($"Unable to delete cache file: {cleanupEx.Message}");
                    }
                }
            }

            // Optional: Verify the merged workbook by loading it
            try
            {
                Workbook mergedWorkbook = new Workbook(outputFile);
                Console.WriteLine($"Merged workbook contains {mergedWorkbook.Worksheets.Count} worksheet(s).");
            }
            catch (Exception verifyEx)
            {
                Console.WriteLine($"Error verifying merged workbook: {verifyEx.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                MergeLargeXlsFiles.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
