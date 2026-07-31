// Title: Prevent saving an Aspose.Cells workbook when worksheets have duplicate TabId values (C#)
// Description: Demonstrates how to detect duplicate TabId values across worksheets in an Aspose.Cells Workbook using a HashSet, and abort the Save operation if any conflict is found. The example creates three sheets, intentionally repeats a TabId, validates uniqueness with a HasDuplicateTabIds method, and shows how to handle the error before calling workbook.Save.
// Keywords: Aspose.Cells | C# | duplicate TabId | worksheet TabId validation | prevent workbook save | TabId conflict detection | HashSet duplicate check | Excel sheet identifier uniqueness
// Common Searches: Aspose.Cells prevent save on duplicate TabId | C# check for duplicate worksheet TabId | validate unique TabId before workbook.Save | detect repeated TabId in Aspose.Cells workbook | how to stop saving Excel file with duplicate sheet IDs
// Developer Intent: Ensure that a workbook is not saved when any worksheet contains a non‑unique TabId, protecting the file from identifier conflicts.
// Use Cases: Integrate HasDuplicateTabIds into a data‑import pipeline that creates worksheets dynamically, aborting the export if TabId collisions occur. | Replace the console warning with an exception to enforce TabId uniqueness in enterprise‑level document generation services. | Run the duplicate‑TabId check after renaming or reordering sheets in an automated report generator to guarantee compliance with downstream processing requirements.
// AI Prompts: Generate a LINQ‑based version of HasDuplicateTabIds that returns the list of duplicate TabId values and their worksheet names. | Write NUnit tests for the duplicate TabId detection logic covering both unique and colliding scenarios. | Create a reusable Aspose.Cells extension method that validates TabId uniqueness and throws a custom DuplicateTabIdException.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace PreventDuplicateTabIdSave
{
    // Demonstrates how to detect duplicate TabId values across worksheets in an Aspose.Cells Workbook using a HashSet, and abort the Save operation if any conflict is found. The example creates three sheets, intentionally repeats a TabId, validates uniqueness with a HasDuplicateTabIds method, and shows how to handle the error before calling workbook.Save.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: use constructor)
            Workbook workbook = new Workbook();

            // Add a few worksheets and intentionally set duplicate TabId values
            int sheet1Idx = workbook.Worksheets.Add();
            Worksheet sheet1 = workbook.Worksheets[sheet1Idx];
            sheet1.Name = "SheetA";
            sheet1.TabId = 101; // first unique TabId

            int sheet2Idx = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Idx];
            sheet2.Name = "SheetB";
            sheet2.TabId = 102; // second unique TabId

            int sheet3Idx = workbook.Worksheets.Add();
            Worksheet sheet3 = workbook.Worksheets[sheet3Idx];
            sheet3.Name = "SheetC";
            sheet3.TabId = 101; // duplicate TabId (same as SheetA)

            // Perform any other modifications here...
            // ...

            // Validate that no duplicate TabId exists before saving
            if (HasDuplicateTabIds(workbook))
            {
                Console.WriteLine("Error: Duplicate TabId detected. Workbook will not be saved.");
                // Optionally, you could throw an exception or handle the situation as needed
                // throw new InvalidOperationException("Duplicate TabId values found.");
            }
            else
            {
                // Save the workbook (lifecycle rule: use Save method)
                string outputPath = "ValidatedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }

            // Clean up
            workbook.Dispose();
        }

        static bool HasDuplicateTabIds(Workbook workbook)
        {
            HashSet<int> seenTabIds = new HashSet<int>();
            foreach (Worksheet ws in workbook.Worksheets)
            {
                int tabId = ws.TabId;
                // If the TabId has already been encountered, we have a duplicate
                if (!seenTabIds.Add(tabId))
                {
                    // Duplicate found
                    Console.WriteLine($"Duplicate TabId {tabId} found in worksheet '{ws.Name}'.");
                    return true;
                }
            }
            // No duplicates
            return false;
        }
    }
}
