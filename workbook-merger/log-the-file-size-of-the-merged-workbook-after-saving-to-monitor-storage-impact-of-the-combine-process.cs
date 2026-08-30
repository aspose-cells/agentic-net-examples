// Title: Merge multiple Excel workbooks with Aspose.Cells for .NET and log the saved file size
// AI Prompts: Write C# code that uses Aspose.Cells to combine several .xlsx files and prints the byte size of the resulting file after saving. | Create a program that generates sample workbooks, merges them with Workbook.Combine, saves the merged workbook, and outputs the file length. | Show how to measure the storage impact of merging Excel files by retrieving the FileInfo.Length after calling Workbook.Save in C#.
// Common Searches: C# Aspose.Cells combine two workbooks and get output file size | how to retrieve size of merged Excel file after using Workbook.Combine | Aspose.Cells .NET log file size after saving merged workbook | measure storage increase when merging multiple .xlsx files with Aspose.Cells | get byte count of saved workbook in C# Aspose.Cells
// Tags: combine workbooks Aspose.Cells C# | retrieve saved workbook size .NET | measure merged Excel file size | Workbook.Combine storage impact | track file length after Aspose.Cells Save | monitor Excel merge file size

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeAndLogSize
{
    // The example creates two sample Excel workbooks, merges them into a single workbook using Aspose.Cells' Combine method, saves the merged file, and then logs the saved file's size in bytes before cleaning up the temporary source files.
    class Program
    {
        static void Main()
        {
            // Prepare temporary files to be merged
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };
            CreateSampleWorkbook(sourceFiles[0], "Data from source 1");
            CreateSampleWorkbook(sourceFiles[1], "Data from source 2");

            // Load the first workbook which will receive the others
            Workbook mergedWorkbook = new Workbook(sourceFiles[0]);

            // Load and combine the remaining workbooks
            for (int i = 1; i < sourceFiles.Length; i++)
            {
                Workbook wb = new Workbook(sourceFiles[i]);
                mergedWorkbook.Combine(wb);
                wb.Dispose();
            }

            // Define output path for the merged workbook
            string mergedFilePath = "MergedResult.xlsx";

            // Save the merged workbook (uses the provided Save method)
            mergedWorkbook.Save(mergedFilePath);
            mergedWorkbook.Dispose();

            // Log the file size after saving
            FileInfo info = new FileInfo(mergedFilePath);
            Console.WriteLine($"Merged workbook saved to '{mergedFilePath}'.");
            Console.WriteLine($"File size: {info.Length} bytes.");

            // Clean up temporary source files
            foreach (string file in sourceFiles)
            {
                if (File.Exists(file))
                    File.Delete(file);
            }
        }

        // Helper to create a simple workbook with a single cell value and save it
        private static void CreateSampleWorkbook(string filePath, string cellValue)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue(cellValue);
            wb.Save(filePath);
            wb.Dispose();
        }
    }
}
