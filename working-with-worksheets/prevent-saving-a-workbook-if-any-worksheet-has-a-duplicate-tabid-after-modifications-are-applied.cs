// Title: Prevent Saving an Aspose.Cells Workbook When Worksheets Have Duplicate TabId Values (C#)
// Description: Demonstrates how to scan all worksheets in an Aspose.Cells workbook for duplicate TabId values using a HashSet, abort the save operation if a conflict is found, and optionally correct the IDs before calling Workbook.Save. The example creates three sheets, deliberately duplicates a TabId, and shows the validation logic.
// Keywords: Aspose.Cells duplicate TabId | C# worksheet TabId uniqueness | prevent workbook save Aspose.Cells | TabId validation .NET | Aspose.Cells Workbook.Save check | detect duplicate worksheet IDs | HashSet duplicate detection C#
// Common Searches: Aspose.Cells check for duplicate TabId before saving | C# prevent workbook save if TabId conflict | validate worksheet TabId uniqueness Aspose.Cells | how to detect duplicate TabId in Aspose.Cells workbook | Aspose.Cells TabId duplicate handling
// Developer Intent: Ensure a workbook is not saved when any worksheet contains a non‑unique TabId.
// Use Cases: Validate TabId uniqueness after adding, removing, or renaming sheets in a batch generation process. | Automatically assign new unique TabIds to colliding worksheets before invoking Workbook.Save. | Integrate duplicate TabId detection into CI pipelines that produce Excel files with Aspose.Cells.
// AI Prompts: Generate a method that returns all duplicate TabId values in an Aspose.Cells workbook. | Show code to reassign unique TabIds to worksheets that share the same TabId before saving. | Create a custom DuplicateTabIdException and demonstrate throwing it when duplicate TabIds are detected during save.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to scan all worksheets in an Aspose.Cells workbook for duplicate TabId values using a HashSet, abort the save operation if a conflict is found, and optionally correct the IDs before calling Workbook.Save. The example creates three sheets, deliberately duplicates a TabId, and shows the validation logic.
class PreventDuplicateTabIdSave
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the default worksheet and set a TabId
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "Sheet1";
        ws1.TabId = 101;

        // Add a second worksheet with a different TabId
        int idx2 = workbook.Worksheets.Add();
        Worksheet ws2 = workbook.Worksheets[idx2];
        ws2.Name = "Sheet2";
        ws2.TabId = 102;

        // Add a third worksheet that intentionally duplicates a TabId
        int idx3 = workbook.Worksheets.Add();
        Worksheet ws3 = workbook.Worksheets[idx3];
        ws3.Name = "Sheet3";
        ws3.TabId = 101; // Duplicate of ws1

        // Validate TabId uniqueness before attempting to save
        if (HasDuplicateTabIds(workbook))
        {
            Console.WriteLine("Cannot save workbook: duplicate TabId detected.");
            // Handle the situation as needed (e.g., assign new TabIds, abort, etc.)
        }
        else
        {
            // No duplicates – safe to save
            workbook.Save("ValidWorkbook.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }

        // Release resources
        workbook.Dispose();
    }

    // Checks the workbook for duplicate TabId values.
    // Returns true if any duplicate is found.
    static bool HasDuplicateTabIds(Workbook wb)
    {
        HashSet<int> seenIds = new HashSet<int>();
        foreach (Worksheet ws in wb.Worksheets)
        {
            int currentId = ws.TabId;
            if (seenIds.Contains(currentId))
            {
                // Duplicate found
                return true;
            }
            seenIds.Add(currentId);
        }
        // All TabIds are unique
        return false;
    }
}
