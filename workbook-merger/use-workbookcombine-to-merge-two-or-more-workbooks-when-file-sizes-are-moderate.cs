// Title: Merge several moderate-size Excel workbooks into one file with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells merge multiple .xlsx workbooks into a single workbook | How to combine worksheets from different Excel files using Aspose.Cells for .NET | Combine three Excel files programmatically with Aspose.Cells when file size is moderate | Aspose.Cells example to copy worksheets from several workbooks into one workbook | Merge Excel workbooks without exceeding memory limits using Aspose.Cells C#
// Tags: Aspose.Cells workbook worksheet copy merging | C# combine multiple .xlsx files Aspose.Cells | merge Excel workbooks moderate size Aspose.Cells | addcopy method workbook merging Aspose.Cells | programmatic Excel workbook consolidation .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies each source file, loads the first workbook as the base, iterates through the remaining workbooks, copies each worksheet into the base workbook using AddCopy, and saves the combined result as a new .xlsx file while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Paths of the workbooks to be merged (moderate file sizes)
            string[] sourceFiles = { "Book1.xlsx", "Book2.xlsx", "Book3.xlsx" };

            // Path for the combined workbook
            string destinationFile = "Combined.xlsx";

            // Verify that all source files exist
            foreach (string file in sourceFiles)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"Source file not found: {file}");
                    return;
                }
            }

            // Load the first workbook as the base for merging
            Workbook combinedWorkbook = new Workbook(sourceFiles[0]);

            // Append worksheets from the remaining workbooks
            for (int i = 1; i < sourceFiles.Length; i++)
            {
                Workbook wb = new Workbook(sourceFiles[i]);
                foreach (Worksheet ws in wb.Worksheets)
                {
                    // Copy each worksheet into the combined workbook by name
                    combinedWorkbook.Worksheets.AddCopy(ws.Name);
                }
            }

            // Save the merged workbook
            combinedWorkbook.Save(destinationFile);
            Console.WriteLine($"Workbooks combined successfully into '{destinationFile}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
