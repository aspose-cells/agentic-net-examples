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
                @"C:\Data\LargeFile1.xls",
                @"C:\Data\LargeFile2.xls",
                // Add more file paths as needed
            };

            // Temporary cache file required by MergeFiles
            string cacheFile = Path.Combine(Path.GetTempPath(), "MergeCache.tmp");

            // Destination merged file
            string mergedFile = @"C:\Data\MergedOutput.xls";

            try
            {
                // Merge the specified files into a single workbook
                CellsHelper.MergeFiles(filesToMerge, cacheFile, mergedFile);
                Console.WriteLine($"Files merged successfully. Output saved to: {mergedFile}");

                // Optional: Load the merged workbook to verify the result
                Workbook mergedWorkbook = new Workbook(mergedFile);
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
                    try { File.Delete(cacheFile); } catch { /* ignore cleanup errors */ }
                }
            }
        }
    }
}