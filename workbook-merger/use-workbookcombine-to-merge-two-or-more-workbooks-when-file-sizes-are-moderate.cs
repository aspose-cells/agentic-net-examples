// Title: Merge Multiple Excel Workbooks with Aspose.Cells Workbook.Combine in C#
// Description: Shows how to create three Workbook objects, add sample data, use the first workbook as the base, merge the other workbooks with the Workbook.Combine method, and save the combined result as an XLSX file.
// Keywords: Aspose.Cells | Workbook.Combine | C# merge Excel workbooks | combine multiple XLSX files .NET | moderate size workbook merge | Aspose.Cells merge workbooks
// Common Searches: Aspose.Cells combine workbooks C# | How to merge several Excel files using Workbook.Combine | Combine Excel workbooks without high memory usage Aspose | C# code to merge multiple workbooks into one | Workbook.Combine method example
// Developer Intent: Programmatically merge two or more Excel workbooks into a single workbook using Aspose.Cells for .NET.
// Use Cases: Consolidate monthly reports into a master workbook for executive review. | Aggregate department‑specific worksheets into one file for cross‑team analysis. | Attach a standard template to user‑generated workbooks before final distribution. | Combine data‑export files from separate systems into a single spreadsheet for import.
// AI Prompts: Write C# code that loops through a list of Workbook objects and merges each into a base workbook using Workbook.Combine, with comprehensive error handling. | Provide an example that merges workbooks while preserving original worksheet names and automatically renaming any duplicates. | Show how to combine multiple workbooks, apply a uniform style to all resulting worksheets, and then save the file as XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create three Workbook objects, add sample data, use the first workbook as the base, merge the other workbooks with the Workbook.Combine method, and save the combined result as an XLSX file.
    public class WorkbookCombineDemo
    {
        public static void Run()
        {
            try
            {
                // Create the first workbook and add some data
                Workbook wb1 = new Workbook();
                wb1.Worksheets[0].Cells["A1"].PutValue("Workbook 1 - Data");

                // Create the second workbook and add some data
                Workbook wb2 = new Workbook(FileFormatType.Xlsx);
                wb2.Worksheets[0].Cells["B2"].PutValue("Workbook 2 - Data");

                // Create the third workbook and add some data
                Workbook wb3 = new Workbook();
                wb3.Worksheets[0].Cells["C3"].PutValue("Workbook 3 - Data");

                // Use the first workbook as the base for merging
                Workbook combinedWorkbook = wb1;

                // Merge the remaining workbooks into the base workbook
                combinedWorkbook.Combine(wb2);
                combinedWorkbook.Combine(wb3);

                // Save the merged workbook to disk
                combinedWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookCombineDemo.Run();
        }
    }
}
