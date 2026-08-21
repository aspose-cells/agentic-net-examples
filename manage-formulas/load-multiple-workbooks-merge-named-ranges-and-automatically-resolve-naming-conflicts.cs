// Title: C# – Merge Multiple Excel Workbooks and Consolidate Named Ranges with Automatic Conflict Resolution using Aspose.Cells
// Description: Loads a primary workbook, iterates through additional .xlsx files, copies each named range, detects case‑insensitive name collisions, appends a numeric suffix to create unique names, removes any duplicate entries, and saves the combined workbook.
// Keywords: Aspose.Cells | C# merge workbooks | named range conflict resolution | automatic rename named ranges | remove duplicate names | Excel workbook consolidation | Aspose.Cells .NET example | GitHub Aspose.Cells code
// Common Searches: Aspose.Cells merge workbooks C# | combine named ranges from multiple Excel files | resolve duplicate named ranges Aspose.Cells | C# code to rename conflicting named ranges | remove duplicate named ranges after merging Excel workbooks
// Developer Intent: Programmatically combine several Excel files into a single workbook while preserving all named ranges and ensuring each name is unique.
// Use Cases: Aggregate departmental financial models into a master workbook without name collisions. | Build a reporting package that pulls data from multiple source workbooks and automatically handles duplicate named range identifiers. | Create a distribution workbook that merges template files and their named ranges, guaranteeing unique names for downstream processing.
// AI Prompts: Generate C# code using Aspose.Cells that merges named ranges from a list of workbooks and automatically renames duplicates with a numeric suffix. | Explain the purpose and optimal placement of the Names.RemoveDuplicateNames method in an Aspose.Cells merge workflow. | Provide a step‑by‑step tutorial to load three Excel workbooks, combine their named ranges, resolve naming conflicts, and save the result as a new file with Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace MergeNamedRangesDemo
{
    // Loads a primary workbook, iterates through additional .xlsx files, copies each named range, detects case‑insensitive name collisions, appends a numeric suffix to create unique names, removes any duplicate entries, and saves the combined workbook.
    class Program
    {
        static void Main()
        {
            // Paths of workbooks to be merged
            string[] workbookFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Load the first workbook as the target workbook
            Workbook targetWorkbook = new Workbook(workbookFiles[0]);

            // Keep a set of existing named range texts for quick conflict detection
            HashSet<string> existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Name n in targetWorkbook.Worksheets.Names)
            {
                existingNames.Add(n.Text);
            }

            // Process remaining workbooks
            for (int i = 1; i < workbookFiles.Length; i++)
            {
                // Load source workbook
                Workbook sourceWorkbook = new Workbook(workbookFiles[i]);

                // Iterate through each named range in the source workbook
                foreach (Name srcName in sourceWorkbook.Worksheets.Names)
                {
                    string srcText = srcName.Text; // The name identifier
                    string srcRefersTo = srcName.RefersTo; // The range reference (e.g., "=Sheet1!$A$1:$B$2")

                    // Resolve naming conflict by generating a unique name if needed
                    string finalName = srcText;
                    int suffix = 1;
                    while (existingNames.Contains(finalName))
                    {
                        finalName = $"{srcText}_{suffix}";
                        suffix++;
                    }

                    // Add the (possibly renamed) named range to the target workbook
                    int idx = targetWorkbook.Worksheets.Names.Add(finalName);
                    Name newName = targetWorkbook.Worksheets.Names[idx];
                    newName.RefersTo = srcRefersTo;

                    // Record the new name to prevent future conflicts
                    existingNames.Add(finalName);
                }

                // Optional: clean any accidental duplicates (safety net)
                targetWorkbook.Worksheets.Names.RemoveDuplicateNames();
            }

            // Save the merged workbook
            string outputPath = "MergedWorkbook_WithNamedRanges.xlsx";
            targetWorkbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }
    }
}
