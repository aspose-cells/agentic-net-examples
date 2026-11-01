// Title: Combine several large XLS workbooks into one file with Aspose.Cells CellsHelper.MergeFiles in C#
// AI Prompts: Generate C# code that accepts a list of .xls file locations, verifies each path, ensures the destination folder exists, and invokes Aspose.Cells.CellsHelper.MergeFiles to create a single merged workbook while handling possible exceptions. | Provide a C# example showing how to combine multiple large Excel 97‑2003 files into one .xls output using Aspose.Cells, including pre‑merge validation and preparation of the output directory.
// Common Searches: c# how to combine many .xls spreadsheets into a single workbook using Aspose.Cells | using CellsHelper.MergeFiles to merge large Excel files in .net core | sample program for merging multiple Excel 97-2003 files with Aspose.Cells C# | check file existence before merging Excel workbooks with Aspose.Cells
// Tags: aspose.cells merge xls workbooks c# | cellshelper.mergefiles example | pre-merge file existence check c# | output directory creation for excel merge | exception handling aspnet excel merging

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    // // Checks each input .xls file, creates the output folder if missing, and merges the workbooks into a single .xls file using Aspose.Cells.CellsHelper.MergeFiles with exception handling.
    class Program
    {
        static void Main()
        {
            // Input workbook files (XLS or XLSX)
            string[] inputFiles = new string[]
            {
                @"C:\Data\File1.xls",
                @"C:\Data\File2.xls",
                @"C:\Data\File3.xls"
                // Add more file paths as needed
            };

            // Output merged workbook
            string outputFile = @"C:\Data\MergedOutput.xls";

            // Verify that all input files exist
            foreach (string file in inputFiles)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"Input file not found: {file}");
                    return;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // Merge the workbooks; third parameter is password (empty for none)
                CellsHelper.MergeFiles(inputFiles, outputFile, string.Empty);
                Console.WriteLine($"Files merged successfully into: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while merging files: {ex.Message}");
            }
        }
    }
}
