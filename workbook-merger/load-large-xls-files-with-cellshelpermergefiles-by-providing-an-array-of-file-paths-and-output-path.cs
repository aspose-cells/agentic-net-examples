// Title: Combine several large XLS workbooks into one file using Aspose.Cells CellsHelper.MergeFiles in C#
// Common Searches: how to merge multiple .xls files into a single workbook with Aspose.Cells in C# | using CellsHelper.MergeFiles to concatenate large Excel files without loading them into memory | C# example for merging three XLS files into one output workbook using Aspose.Cells | merge large XLS workbooks programmatically while preserving passwords with CellsHelper | Aspose.Cells merge files array of paths to create a combined XLS workbook in .NET
// Tags: CellsHelper.MergeFiles XLS workbook merging | C# Aspose.Cells batch XLS file concatenation | merge multiple Excel files without loading into memory | programmatic merging of large .xls workbooks | Aspose.Cells combine source files array into single workbook

using System;
using System.IO;
using Aspose.Cells;

// The sample defines an array of source .xls file paths, validates each file, ensures the output directory exists, and then merges all workbooks into a single destination file using CellsHelper.MergeFiles with no password, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        // Array of source XLS file paths to be merged
        string[] sourceFiles = new string[]
        {
            @"C:\Data\File1.xls",
            @"C:\Data\File2.xls",
            @"C:\Data\File3.xls"
        };

        // Destination path for the merged workbook
        string outputFile = @"C:\Data\MergedOutput.xls";

        try
        {
            // Verify that each source file exists
            foreach (string file in sourceFiles)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"Source file not found: {file}");
                    return;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Merge all source workbooks into a single workbook.
            // The third parameter is the password for the destination file (null if none).
            CellsHelper.MergeFiles(sourceFiles, outputFile, null);

            Console.WriteLine("Merge completed successfully. Output file: " + outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during merging: " + ex.Message);
        }
    }
}
