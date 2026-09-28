// Title: Add Aspose.Cells via NuGet and merge multiple .xlsx workbooks into a single file using CellsHelper.MergeFiles in C#
// AI Prompts: Generate C# code that installs the Aspose.Cells NuGet package, creates two sample .xlsx workbooks, and merges them into one workbook using CellsHelper.MergeFiles with a temporary cache file. | Write a .NET console application that checks that source Excel files exist, merges them into a destination workbook via CellsHelper.MergeFiles, and then deletes the source and cache files while handling any exceptions. | Provide a step‑by‑step guide for adding the Aspose.Cells reference through NuGet, configuring a temporary cache for the merge operation, and performing workbook merging with proper error handling in C#.
// Common Searches: aspnet add Aspose.Cells NuGet package then merge two xlsx files programmatically | c# CellsHelper.MergeFiles example with temporary cache file | how to combine multiple Excel workbooks into one using Aspose.Cells in a console app | merge Excel files and delete source files after merge using Aspose.Cells C#
// Tags: install Aspose.Cells NuGet package for .NET | cellshelper.mergefiles combine .xlsx workbooks | temporary cache file usage in Aspose.Cells merge | exception handling for workbook merging in C# | cleanup source and cache files after Excel merge

using System;
using System.IO;
using Aspose.Cells;

namespace MergeExcelFilesDemo
{
    // This C# console example shows how to add the Aspose.Cells library via NuGet, create two sample .xlsx workbooks, and merge them into a single workbook using CellsHelper.MergeFiles with a temporary cache file. It includes validation of source files, error handling, and cleanup of the original files and cache after the merge.
    class Program
    {
        static void Main(string[] args)
        {
            // Prepare sample workbooks to be merged
            string[] sourceFiles = new string[] { "Source1.xlsx", "Source2.xlsx" };

            try
            {
                // Create first sample workbook
                Workbook wb1 = new Workbook();
                wb1.Worksheets[0].Cells["A1"].PutValue("Data from first file");
                wb1.Save(sourceFiles[0]);

                // Create second sample workbook
                Workbook wb2 = new Workbook();
                wb2.Worksheets[0].Cells["A1"].PutValue("Data from second file");
                wb2.Save(sourceFiles[1]);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating sample workbooks: {ex.Message}");
                return;
            }

            // Define temporary cache file required by CellsHelper.MergeFiles
            string cacheFile = "mergeCache.tmp";

            // Define destination file for the merged result
            string destFile = "MergedResult.xlsx";

            try
            {
                // Ensure source files exist before merging
                foreach (var file in sourceFiles)
                {
                    if (!File.Exists(file))
                        throw new FileNotFoundException($"Source file not found: {file}");
                }

                // Merge the source Excel files into a single workbook
                CellsHelper.MergeFiles(sourceFiles, cacheFile, destFile);

                Console.WriteLine($"Files merged successfully. Output saved to '{destFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merge: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files
                foreach (var file in sourceFiles)
                {
                    if (File.Exists(file))
                        File.Delete(file);
                }

                if (File.Exists(cacheFile))
                    File.Delete(cacheFile);
            }
        }
    }
}
