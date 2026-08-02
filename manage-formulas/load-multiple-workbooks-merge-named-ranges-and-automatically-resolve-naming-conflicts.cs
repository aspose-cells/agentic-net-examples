// Title: Merge Excel workbooks & consolidate named ranges with conflict resolution – Aspose.Cells for .NET (C#)
// Description: Loads a set of Excel files, uses the first as the target workbook, and iteratively combines the rest with Workbook.Combine. For each source workbook the code copies its named ranges, detects name collisions and appends a unique suffix, then removes any duplicate definitions before saving the merged file.
// Keywords: Aspose.Cells | C# Excel merge | Workbook.Combine | named ranges | conflict resolution | duplicate name removal | Excel consolidation | merge worksheets | Aspose.Cells API | Save as Xlsx
// Common Searches: Aspose.Cells merge workbooks C# | Combine Excel files and keep named ranges unique | Rename duplicate named ranges when merging workbooks | How to remove duplicate named range definitions Aspose.Cells | C# code to consolidate multiple Excel workbooks
// Developer Intent: Create a single workbook that contains all sheets, data, styles, and uniquely merged named ranges from several source files.
// Use Cases: Monthly financial reports from different departments combined into a master workbook with distinct named ranges. | Data‑analysis pipeline that aggregates separate model files while preserving each range's reference. | Distribution package that bundles regional spreadsheets into one file, automatically renaming overlapping names.
// AI Prompts: Generate C# code using Aspose.Cells to merge several workbooks and automatically rename colliding named ranges. | Show how to use Workbook.Combine together with Names.RemoveDuplicateNames to produce a clean merged Excel file. | Explain the algorithm for suffix‑based conflict resolution of named ranges during workbook consolidation.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsMergeNamedRanges
{
    // Loads a set of Excel files, uses the first as the target workbook, and iteratively combines the rest with Workbook.Combine. For each source workbook the code copies its named ranges, detects name collisions and appends a unique suffix, then removes any duplicate definitions before saving the merged file.
    class Program
    {
        static void Main()
        {
            // Paths of workbooks to be merged
            string[] files = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Load the first workbook – it will become the target workbook
            Workbook target = new Workbook(files[0]);

            // Process remaining workbooks
            for (int i = 1; i < files.Length; i++)
            {
                // Load source workbook
                Workbook source = new Workbook(files[i]);

                // Combine worksheets, data, styles, etc.
                target.Combine(source);

                // Merge named ranges from source into target
                MergeNamedRanges(target, source, i);
            }

            // Remove any duplicate name definitions that might have been created
            target.Worksheets.Names.RemoveDuplicateNames();

            // Save the merged workbook
            string outputPath = "MergedResult.xlsx";
            target.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }

        /// <param name="target">The workbook that will receive the named ranges.</param>
        /// <param name="source">The workbook providing named ranges.</param>
        /// <param name="sourceIndex">Zero‑based index of the source workbook (used for generating unique names).</param>
        private static void MergeNamedRanges(Workbook target, Workbook source, int sourceIndex)
        {
            // Build a quick lookup of existing names in the target workbook
            HashSet<string> existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Name tn in target.Worksheets.Names)
            {
                existingNames.Add(tn.Text);
            }

            // Iterate through each named range in the source workbook
            foreach (Name srcName in source.Worksheets.Names)
            {
                string newName = srcName.Text;

                // Resolve conflict by appending a suffix until the name becomes unique
                if (existingNames.Contains(newName))
                {
                    int suffix = 1;
                    string baseName = newName;
                    do
                    {
                        newName = $"{baseName}_From{sourceIndex}_{suffix}";
                        suffix++;
                    } while (existingNames.Contains(newName));
                }

                // Add the (possibly renamed) name to the target workbook
                int idx = target.Worksheets.Names.Add(newName);
                Name addedName = target.Worksheets.Names[idx];
                addedName.RefersTo = srcName.RefersTo; // copy the reference (e.g., =Sheet1!$A$1:$B$2)

                // Keep the lookup up‑to‑date
                existingNames.Add(newName);
            }
        }
    }
}
