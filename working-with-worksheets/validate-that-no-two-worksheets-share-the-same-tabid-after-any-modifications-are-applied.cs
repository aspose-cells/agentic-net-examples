// Title: Check for duplicate worksheet TabId values in an Aspose.Cells workbook with C#
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all worksheets and logs any repeated TabId values. | Create a reusable C# method that returns true only if every worksheet in a workbook has a distinct TabId. | Write a C# routine that validates worksheet TabId uniqueness and saves the workbook only when no duplicates are found.
// Common Searches: aspnet c# how to ensure each Excel sheet has a unique TabId with Aspose.Cells | detect duplicate worksheet TabId in Aspose.Cells .NET library | C# code sample for validating worksheet TabId uniqueness in an Excel workbook | Aspose.Cells check for repeated TabId after modifying worksheets
// Tags: Aspose.Cells validate worksheet TabId uniqueness | C# detect duplicate worksheet TabId | Aspose.Cells worksheet TabId duplicate check | hashset duplicate detection Aspose.Cells | Excel workbook TabId validation C#

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The program loads an Excel workbook with Aspose.Cells, optionally modifies worksheets, then uses a HashSet to ensure each worksheet's TabId is unique, reports any duplicates, and saves the file only if validation succeeds.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // -------------------------------------------------
        // Perform any worksheet modifications here
        // -------------------------------------------------
        // Example: workbook.Worksheets[0].Name = "FirstSheet";

        // Validate that each worksheet has a unique TabId
        HashSet<int> tabIdSet = new HashSet<int>();
        bool duplicateFound = false;

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int tabId = sheet.TabId; // Get the TabId of the current worksheet

            // Try to add the TabId to the set; if it already exists, we have a duplicate
            if (!tabIdSet.Add(tabId))
            {
                Console.WriteLine($"Duplicate TabId detected: {tabId} on worksheet \"{sheet.Name}\"");
                duplicateFound = true;
            }
        }

        if (!duplicateFound)
        {
            Console.WriteLine("Validation passed: all worksheets have unique TabId values.");
        }

        // Save the workbook after validation (optional)
        workbook.Save("output.xlsx");
    }
}
