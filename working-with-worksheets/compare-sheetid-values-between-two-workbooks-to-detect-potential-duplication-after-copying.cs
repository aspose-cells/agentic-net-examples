// Title: Detect Duplicate Worksheet TabId After Workbook.Copy with Aspose.Cells for .NET
// Description: Loads a source and a destination workbook, copies all worksheets from the source into the destination using Workbook.Copy, gathers each source worksheet's TabId, scans the merged workbook for matching TabId values, reports any duplicates, and saves the result as a new file.
// Keywords: Aspose.Cells TabId duplicate detection | compare worksheet IDs .NET | Workbook.Copy duplicate sheets | Aspose.Cells merge workbook check | C# detect duplicate sheet identifiers
// Common Searches: Aspose.Cells find duplicate worksheet TabId after copy | how to compare sheet IDs between two Excel files in C# | detect duplicate sheets after Workbook.Copy Aspose.Cells | check for conflicting TabId values when merging workbooks | C# Aspose.Cells duplicate sheet identifier example
// Developer Intent: Identify worksheets that share the same TabId after merging two workbooks to prevent identifier conflicts.
// Use Cases: Validate that merging workbooks does not create duplicate TabId values before saving. | Log duplicate sheet names and IDs to troubleshoot Excel merging operations. | Ensure unique worksheet identifiers when programmatically consolidating multiple Excel files.
// AI Prompts: Write C# code with Aspose.Cells that copies a workbook and then lists any worksheets whose TabId already exists in the source workbook. | Show an Aspose.Cells .NET example that verifies unique TabId values after Workbook.Copy and handles duplicates gracefully. | Explain how Aspose.Cells assigns TabId to worksheets and how to renumber or reset them to avoid duplication after a merge.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SheetIdComparisonDemo
{
    // Loads a source and a destination workbook, copies all worksheets from the source into the destination using Workbook.Copy, gathers each source worksheet's TabId, scans the merged workbook for matching TabId values, reports any duplicates, and saves the result as a new file.
    class Program
    {
        static void Main()
        {
            // Paths to the workbooks
            string sourcePath = "source.xlsx";
            string destinationPath = "destination.xlsx";

            // Load the source and destination workbooks (using the provided constructor rule)
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Workbook destinationWorkbook = new Workbook(destinationPath);

            // Copy the entire source workbook into the destination workbook
            // (uses the Workbook.Copy(Workbook) rule)
            destinationWorkbook.Copy(sourceWorkbook);

            // Collect TabId values from the source workbook
            HashSet<int> sourceTabIds = new HashSet<int>();
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                sourceTabIds.Add(ws.TabId);
            }

            // Compare TabId values in the destination workbook against the source set
            Console.WriteLine("Checking for duplicate Sheet (TabId) values after copy:");
            foreach (Worksheet ws in destinationWorkbook.Worksheets)
            {
                if (sourceTabIds.Contains(ws.TabId))
                {
                    Console.WriteLine($"Duplicate found - Sheet Name: \"{ws.Name}\", TabId: {ws.TabId}");
                }
            }

            // Save the merged workbook (using the provided Save method)
            destinationWorkbook.Save("merged_output.xlsx");
        }
    }
}
