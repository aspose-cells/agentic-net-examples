// Title: Combine multiple Excel workbooks into a single XLSX file using Aspose.Cells Workbook.Combine in C#
// AI Prompts: Write C# code that creates a destination Workbook, loads two existing Excel files, merges them with Workbook.Combine, and saves the result as CombinedWorkbook.xlsx. | Show how to wrap Workbook.Combine calls in try‑catch blocks to handle errors while merging several workbooks and ensure the final file is saved in XLSX format. | Modify the example to accept an array of file paths at runtime and combine all referenced workbooks into one using Aspose.Cells Workbook.Combine.
// Common Searches: Aspose.Cells C# how to merge three workbooks into one XLSX | using Workbook.Combine to concatenate Excel files with moderate size | C# example for combining in-memory workbooks with Aspose.Cells | merge multiple Excel workbooks programmatically with Aspose.Cells Workbook.Combine | error handling when combining workbooks using Aspose.Cells C#
// Tags: Aspose.Cells Workbook.Combine for XLSX output | C# programmatic Excel workbook merging | merge in-memory workbooks with Aspose.Cells | combine multiple workbooks without exceeding memory limits | save merged workbook as XLSX format

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a destination workbook, adds two additional workbooks in memory, merges them using Workbook.Combine, and saves the combined result as CombinedWorkbook.xlsx.
    public class WorkbookCombineDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create the first workbook (will become the destination)
            Workbook combinedWorkbook = new Workbook(FileFormatType.Xlsx);
            combinedWorkbook.Worksheets[0].Cells["A1"].PutValue("Workbook 1");

            // Create additional workbooks to be merged
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Workbook 2");

            Workbook wb3 = new Workbook();
            wb3.Worksheets[0].Cells["A1"].PutValue("Workbook 3");

            // Merge the second workbook into the combined workbook
            combinedWorkbook.Combine(wb2);

            // Merge the third workbook into the combined workbook
            combinedWorkbook.Combine(wb3);

            // Save the final merged workbook
            combinedWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
