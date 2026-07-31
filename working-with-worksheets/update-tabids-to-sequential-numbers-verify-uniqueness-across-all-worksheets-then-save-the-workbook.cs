// Title: Set Sequential Worksheet TabIds and Ensure Uniqueness with Aspose.Cells for .NET
// Description: C# example that creates or loads a workbook, removes the default sheet, adds worksheets, assigns each worksheet a sequential TabId starting at 1, validates that all TabIds are unique using a HashSet, and saves the workbook to disk.
// Keywords: Aspose.Cells TabId | C# worksheet TabId sequential | validate unique TabId Aspose | update worksheet TabId .NET | save workbook after TabId change
// Common Searches: Aspose.Cells assign sequential TabId to worksheets | how to check duplicate TabId in Aspose.Cells workbook | C# set worksheet TabId property Aspose | ensure unique TabId values before saving Excel file | Aspose.Cells TabId validation example
// Developer Intent: Assign a unique, incrementing TabId to every worksheet in a workbook and verify that no duplicates exist before saving.
// Use Cases: Generate a new workbook with multiple sheets, give each sheet a distinct TabId, and export the file. | Load an existing Excel file, re‑order or rename sheets, then renumber TabIds to maintain a clean tab order. | Detect and prevent duplicate TabId values during batch processing of workbooks to avoid Excel UI inconsistencies.
// AI Prompts: Write C# code using Aspose.Cells that iterates through all worksheets, sets TabId = 1,2,3…, checks for duplicates, and saves the workbook. | Create error‑handling logic that throws an exception when a duplicate TabId is found while updating worksheets with Aspose.Cells. | Provide a reusable method: bool UpdateTabIds(Workbook wb) that assigns sequential TabIds, validates uniqueness, and returns true on success.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// C# example that creates or loads a workbook, removes the default sheet, adds worksheets, assigns each worksheet a sequential TabId starting at 1, validates that all TabIds are unique using a HashSet, and saves the workbook to disk.
class UpdateTabIds
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Remove the default worksheet to avoid duplicate names
            if (workbook.Worksheets.Count > 0)
            {
                workbook.Worksheets.RemoveAt(0);
            }

            // Add sample worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Update TabId values to sequential numbers (starting from 1)
            int nextId = 1;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.TabId = nextId++;
            }

            // Verify that all TabId values are unique across the workbook
            HashSet<int> seenIds = new HashSet<int>();
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (!seenIds.Add(ws.TabId))
                {
                    throw new InvalidOperationException($"Duplicate TabId detected: {ws.TabId}");
                }
            }

            // Save the workbook
            string outputPath = "UpdatedTabIds.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
