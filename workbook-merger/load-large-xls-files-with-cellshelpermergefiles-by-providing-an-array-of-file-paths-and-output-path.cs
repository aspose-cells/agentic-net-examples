// Title: C# – Merge Large XLS Workbooks with Aspose.Cells CellsHelper.MergeFiles
// Description: A complete C# example that merges an array of large .xls files into a single workbook using Aspose.Cells CellsHelper.MergeFiles. The code shows how to specify a temporary cache file for memory‑efficient processing, verify the merged worksheet count, and clean up the cache file after the operation.
// Keywords: Aspose.Cells merge large xls | CellsHelper.MergeFiles C# | C# merge .xls files | Excel workbook merger with cache | temporary cache file Aspose | large Excel files consolidation | GitHub Aspose.Cells example | C# Excel file aggregation
// Common Searches: how to merge large xls files in C# with Aspose.Cells | CellsHelper.MergeFiles cache file purpose | C# example for merging multiple .xls workbooks | Aspose.Cells merge files memory efficient | verify worksheet count after merging Excel files
// Developer Intent: Combine several large .xls workbooks into one file while controlling memory usage through a temporary cache.
// Use Cases: Consolidate quarterly departmental reports stored as legacy .xls files into a single workbook for executive review. | Aggregate exported data from multiple legacy systems before performing analysis in a unified Excel file. | Create an archive workbook that merges legacy .xls documents to simplify migration to newer Excel formats.
// AI Prompts: Generate a C# snippet that merges an array of .xls files with CellsHelper.MergeFiles, includes try‑catch error handling, and deletes the temporary cache file. | Explain the role of the temporary cache file in CellsHelper.MergeFiles and how to configure its path for large workbook merges. | Show how to load the merged workbook after CellsHelper.MergeFiles, read the worksheet count, and handle possible exceptions.

using System;
using System.IO;
using Aspose.Cells;

namespace MergeLargeXlsFiles
{
    // A complete C# example that merges an array of large .xls files into a single workbook using Aspose.Cells CellsHelper.MergeFiles. The code shows how to specify a temporary cache file for memory‑efficient processing, verify the merged worksheet count, and clean up the cache file after the operation.
    class Program
    {
        static void Main()
        {
            // Paths of the large XLS files to be merged
            string[] filesToMerge = new string[]
            {
                "LargeFile1.xls",
                "LargeFile2.xls",
                "LargeFile3.xls"
            };

            // Temporary cache file required by CellsHelper.MergeFiles
            string cacheFile = "MergeCache.tmp";

            // Destination file for the merged result
            string outputFile = "MergedResult.xls";

            try
            {
                // Merge the specified files into a single workbook
                CellsHelper.MergeFiles(filesToMerge, cacheFile, outputFile);
                Console.WriteLine($"Files merged successfully. Output saved to: {outputFile}");

                // Load the merged workbook to verify the operation
                Workbook mergedWorkbook = new Workbook(outputFile);
                Console.WriteLine($"Merged workbook contains {mergedWorkbook.Worksheets.Count} worksheet(s).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merging: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary cache file
                if (File.Exists(cacheFile))
                {
                    File.Delete(cacheFile);
                }
            }
        }
    }
}
