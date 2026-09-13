// Title: Identify duplicate worksheet Index values across two Excel workbooks with Aspose.Cells in C#
// AI Prompts: Generate a C# console application that uses Aspose.Cells to load two .xlsx files, collect the Index of each worksheet from the first workbook, and print the names of any worksheets in the second workbook that share the same Index. | Write .NET code that iterates through the Worksheet.Index property of all sheets in two Excel workbooks and outputs a list of potential duplicate sheet IDs.
// Common Searches: C# Aspose.Cells how to compare worksheet indexes between two workbooks | detect duplicate sheet IDs after copying Excel files using Aspose.Cells | find worksheet index collisions in .NET Excel processing | list worksheets with same Index in two Excel workbooks Aspose.Cells
// Tags: Aspose.Cells compare worksheet indexes | detect duplicate sheet IDs .NET | worksheet index collision detection Excel | C# compare worksheet Index property

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// // Loads two Excel workbooks with Aspose.Cells, gathers the Index values of all worksheets from the first workbook, then scans the second workbook for worksheets whose Index already exists, reporting any matching sheet names as potential duplicates.
class SheetIdComparer
{
    static void Main()
    {
        // Paths to the two workbooks to compare
        string workbookPath1 = "Workbook1.xlsx";
        string workbookPath2 = "Workbook2.xlsx";

        // Ensure the input files exist
        if (!File.Exists(workbookPath1))
        {
            Console.WriteLine($"File not found: {workbookPath1}");
            return;
        }
        if (!File.Exists(workbookPath2))
        {
            Console.WriteLine($"File not found: {workbookPath2}");
            return;
        }

        try
        {
            // Load the workbooks
            Workbook wb1 = new Workbook(workbookPath1);
            Workbook wb2 = new Workbook(workbookPath2);

            // Collect all worksheet indexes from the first workbook
            HashSet<int> sheetIndexesInFirst = new HashSet<int>();
            foreach (Worksheet ws in wb1.Worksheets)
            {
                sheetIndexesInFirst.Add(ws.Index);
            }

            // Find worksheets in the second workbook that have duplicate indexes
            List<string> duplicateSheets = new List<string>();
            foreach (Worksheet ws in wb2.Worksheets)
            {
                if (sheetIndexesInFirst.Contains(ws.Index))
                {
                    duplicateSheets.Add(ws.Name);
                }
            }

            // Report the results
            if (duplicateSheets.Count > 0)
            {
                Console.WriteLine("Potential duplicate worksheet indexes detected in the following sheets of the second workbook:");
                foreach (string sheetName in duplicateSheets)
                {
                    Console.WriteLine("- " + sheetName);
                }
            }
            else
            {
                Console.WriteLine("No duplicate worksheet indexes were found between the two workbooks.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the workbooks: " + ex.Message);
        }
    }
}
