// Title: Merge Multiple Excel Workbooks with Aspose.Cells Workbook.Combine (C#)
// Description: Demonstrates how to create two source workbooks, add data, instantiate an empty XLSX destination workbook, merge the sources using Workbook.Combine, and save the combined file as CombinedWorkbook.xlsx. Ideal for moderate‑size files in .NET applications.
// Keywords: Aspose.Cells | Workbook.Combine | C# | merge Excel workbooks | combine workbooks .NET | moderate size Excel files | combine worksheets programmatically | Aspose.Cells example
// Common Searches: Aspose.Cells combine workbooks C# | How to merge Excel files using Workbook.Combine | Combine several .xlsx files in .NET | Merge Excel workbooks without streams Aspose | Programmatic Excel workbook consolidation C#
// Developer Intent: Programmatically merge two or more Excel workbooks into a single workbook using Aspose.Cells for .NET.
// Use Cases: Build a master report by aggregating data from multiple monthly workbooks. | Consolidate product catalogs stored in separate files into one master workbook. | Automate the creation of a combined financial statement from quarterly worksheets.
// AI Prompts: Generate C# code that loops through a list of workbooks and merges them with Workbook.Combine. | Show error handling and logging for combining workbooks of different formats (XLS, XLSX, CSV) using Aspose.Cells. | Explain how to keep original worksheet names and order when merging multiple workbooks. | Provide performance tips for combining moderate‑size Excel files with Workbook.Combine. | Demonstrate how to insert merged worksheets at a specific index in the destination workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create two source workbooks, add data, instantiate an empty XLSX destination workbook, merge the sources using Workbook.Combine, and save the combined file as CombinedWorkbook.xlsx. Ideal for moderate‑size files in .NET applications.
    public class WorkbookCombineDemo
    {
        // Entry point required for console application
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
            // Create the first source workbook and add some data
            Workbook sourceWorkbook1 = new Workbook();
            sourceWorkbook1.Worksheets[0].Cells["A1"].PutValue("Source Workbook 1");

            // Create the second source workbook and add some data
            Workbook sourceWorkbook2 = new Workbook();
            sourceWorkbook2.Worksheets[0].Cells["A1"].PutValue("Source Workbook 2");

            // Create the destination workbook (empty workbook with XLSX format)
            Workbook destinationWorkbook = new Workbook(FileFormatType.Xlsx);
            destinationWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Workbook");

            // Combine the first source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook1);

            // Combine the second source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook2);

            // Save the combined workbook to disk
            destinationWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Combined workbook saved as CombinedWorkbook.xlsx");
        }
    }
}
