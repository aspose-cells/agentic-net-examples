// Title: Prevent Duplicate Worksheet TabId Values in Aspose.Cells for .NET
// Description: Demonstrates how to assign a unique TabId to each worksheet by checking a HashSet of existing IDs, incrementing the desired value until it is free, and then saving the workbook.
// Keywords: Aspose.Cells TabId uniqueness | C# worksheet TabId duplicate prevention | Aspose.Cells set TabId .NET | HashSet TabId tracking | worksheet identifier collision handling
// Common Searches: Aspose.Cells assign unique TabId | prevent duplicate TabId in Excel worksheets | C# check existing TabIds before setting | Aspose.Cells TabId conflict resolution | how to ensure unique worksheet identifiers
// Developer Intent: Assign a TabId to each worksheet only when it does not clash with any TabId already used in the workbook.
// Use Cases: Automatically generate sequential TabIds for newly added worksheets. | Update a worksheet's TabId based on external data while guaranteeing uniqueness. | Validate and correct TabId collisions before exporting a workbook to Excel.
// AI Prompts: Create a C# method for Aspose.Cells that receives a Worksheet, a desired TabId, and a collection of used TabIds, then sets the next available TabId. | Show an Aspose.Cells example that scans all worksheets, detects duplicate TabIds, and resolves them before saving. | Explain how to reset the HashSet of used TabIds when a workbook is closed or reloaded.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to assign a unique TabId to each worksheet by checking a HashSet of existing IDs, incrementing the desired value until it is free, and then saving the workbook.
    public class WorksheetTabIdUniqueDemo
    {
        // Assigns a TabId to a worksheet only if it does not duplicate existing TabIds.
        private static void SetUniqueTabId(Worksheet sheet, int desiredTabId, HashSet<int> existingTabIds)
        {
            // If the desired TabId is already used, find the next free integer.
            int newTabId = desiredTabId;
            while (existingTabIds.Contains(newTabId))
            {
                newTabId++;
            }

            // Apply the unique TabId and record it.
            sheet.TabId = newTabId;
            existingTabIds.Add(newTabId);
            Console.WriteLine($"Worksheet \"{sheet.Name}\" assigned TabId: {newTabId}");
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Ensure there are multiple worksheets to demonstrate the feature.
                Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
                sheet1.Name = "SheetA";

                Worksheet sheet2 = workbook.Worksheets.Add("SheetB");
                Worksheet sheet3 = workbook.Worksheets.Add("SheetC");

                // Collect existing TabIds (initially empty because none have been set)
                HashSet<int> existingTabIds = new HashSet<int>();

                // Assign TabIds, intentionally using duplicate desired values to test the check.
                SetUniqueTabId(sheet1, 100, existingTabIds); // Expected to set 100
                SetUniqueTabId(sheet2, 100, existingTabIds); // Duplicate, should become 101
                SetUniqueTabId(sheet3, 102, existingTabIds); // Unique, stays 102
                SetUniqueTabId(sheet3, 100, existingTabIds); // Attempt duplicate on same sheet, will become 103

                // Save the workbook (lifecycle rule: save)
                string outputPath = "WorksheetTabIdUniqueDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point required for compilation.
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
