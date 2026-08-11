// Title: Renumber Worksheet TabId Sequentially, Validate Uniqueness, and Save Workbook – Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds worksheets, optionally assigns random TabId values, then reassigns each worksheet's TabId to a sequential index starting at 0, checks for duplicate IDs with a HashSet, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | Worksheet TabId | sequential TabId | unique TabId validation | Workbook save | Excel navigation | TabId reset | code example
// Common Searches: Aspose.Cells set worksheet TabId sequentially | how to ensure unique TabId values in Excel workbook | C# update TabId for all sheets | validate duplicate TabId Aspose.Cells | save workbook after TabId renumbering
// Developer Intent: Reassign each worksheet's TabId to a consecutive number, confirm no duplicates exist, and write the workbook to disk.
// Use Cases: Standardize tab order in generated reports before distribution. | Prevent UI glitches when merging workbooks that contain conflicting TabId values. | Prepare a workbook for programmatic navigation where Excel expects sequential TabId identifiers.
// AI Prompts: Write C# code using Aspose.Cells that renumbers all worksheet TabId properties from 0 upward, throws an exception on duplicates, and saves the workbook. | Create a reusable method that takes a Workbook, an optional start index, reassigns TabId values sequentially, validates uniqueness, and returns the modified workbook. | Generate a unit test in C# that verifies the TabId renumbering logic for a workbook initialized with random TabId values.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTabIdUpdater
{
    // C# example that creates a workbook, adds worksheets, optionally assigns random TabId values, then reassigns each worksheet's TabId to a sequential index starting at 0, checks for duplicate IDs with a HashSet, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a few worksheets for demonstration
            Workbook workbook = new Workbook();
            WorksheetCollection sheets = workbook.Worksheets;

            // Ensure we have at least 5 worksheets
            for (int i = 1; i < 5; i++)
            {
                sheets.Add("Sheet" + (i + 1));
            }

            // OPTIONAL: Assign arbitrary (non‑sequential) TabId values to simulate existing data
            Random rnd = new Random();
            foreach (Worksheet ws in sheets)
            {
                ws.TabId = rnd.Next(100, 200);
            }

            // Update TabId values to sequential numbers starting from 0
            for (int i = 0; i < sheets.Count; i++)
            {
                sheets[i].TabId = i;
            }

            // Verify that all TabId values are unique
            HashSet<int> uniqueIds = new HashSet<int>();
            foreach (Worksheet ws in sheets)
            {
                if (!uniqueIds.Add(ws.TabId))
                {
                    throw new InvalidOperationException($"Duplicate TabId detected: {ws.TabId}");
                }
            }

            // Save the workbook
            string outputPath = "UpdatedTabIds.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
