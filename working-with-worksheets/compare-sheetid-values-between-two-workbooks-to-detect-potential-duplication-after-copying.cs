// Title: Detect duplicate worksheets by comparing SheetId (Index) values across two Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that loads two .xlsx files, extracts each worksheet's Index, and prints the names of sheets whose Index appears in both workbooks. | Create a reusable C# method that takes two Excel file paths and returns a list of worksheet names that share the same SheetId when processed with Aspose.Cells.
// Common Searches: aspnet compare worksheet indexes between two Excel files to find copied sheets | c# Aspose.Cells detect duplicate sheet IDs after workbook copy | how to check if worksheets have the same SheetId in two workbooks using Aspose.Cells | find overlapping worksheet indexes in source and target Excel workbooks with Aspose.Cells .NET
// Tags: compare worksheet indexes Aspose.Cells .NET | detect duplicate SheetId across workbooks | worksheet index collision after copy | Aspose.Cells sheet duplication check

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a source and a target Excel workbook, stores the Index values of all worksheets from the source in a HashSet, then iterates the target workbook's worksheets to flag any whose Index matches, outputting potential duplicate sheet names.
class SheetIdComparer
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string targetPath = "TargetWorkbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Verify that the target file exists
            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Target file not found: {targetPath}");
                return;
            }

            // Load the first workbook
            Workbook wbSource = new Workbook(sourcePath);

            // Load the second workbook (the one that may contain copied sheets)
            Workbook wbTarget = new Workbook(targetPath);

            // Collect all worksheet indexes from the source workbook
            HashSet<int> sourceSheetIndexes = new HashSet<int>();
            foreach (Worksheet ws in wbSource.Worksheets)
            {
                sourceSheetIndexes.Add(ws.Index);
            }

            // Compare worksheet indexes in the target workbook against the source set
            bool duplicationFound = false;
            foreach (Worksheet ws in wbTarget.Worksheets)
            {
                if (sourceSheetIndexes.Contains(ws.Index))
                {
                    duplicationFound = true;
                    Console.WriteLine($"Potential duplicate sheet detected: \"{ws.Name}\" (Index={ws.Index})");
                }
            }

            if (!duplicationFound)
            {
                Console.WriteLine("No duplicate worksheet indexes were found between the two workbooks.");
            }

            // Optionally, save any changes or results (not required for comparison)
            // wbTarget.Save("TargetWorkbook_Analyzed.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
