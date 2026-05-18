using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of the large XLS files to be merged
            string[] filesToMerge = new string[]
            {
                "LargeFile1.xls",
                "LargeFile2.xls",
                // Add more file paths as needed
            };

            // Temporary cache file required by CellsHelper.MergeFiles
            string cacheFile = "MergeCache.tmp";

            // Destination file where the merged result will be saved
            string outputFile = "MergedLargeFiles.xls";

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