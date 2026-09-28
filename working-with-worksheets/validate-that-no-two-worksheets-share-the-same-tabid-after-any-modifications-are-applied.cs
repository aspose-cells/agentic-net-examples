// Title: How to verify that all worksheets in an Aspose.Cells workbook have unique TabId values using C#
// AI Prompts: Write a C# method that takes an Aspose.Cells Workbook, iterates through its worksheets, and throws an InvalidOperationException with the sheet name when a duplicate TabId is found. | Generate C# code that uses a HashSet to detect duplicate TabId values across worksheets in an Aspose.Cells workbook and logs each conflict before saving the file.
// Common Searches: aspocells c# check for duplicate worksheet tabid | ensure each Excel sheet has a unique TabId with Aspose.Cells | detect duplicate TabId in workbook using Aspose.Cells C# | C# Aspose.Cells validate worksheet TabId uniqueness before saving | how to enforce unique TabId values in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells worksheet TabId uniqueness check | C# HashSet duplicate TabId detection | Aspose.Cells workbook validation for TabId conflicts | exception handling duplicate TabId Aspose.Cells | unique TabId enforcement in Excel workbook C#

using Aspose.Cells;
using System;
using System.Collections.Generic;

// C# example that iterates through all worksheets in an Aspose.Cells workbook, uses a HashSet to ensure each TabId is unique, throws an InvalidOperationException on duplicates, and saves the workbook if validation passes.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Validate that each worksheet has a unique TabId
        HashSet<int> tabIdSet = new HashSet<int>();
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int tabId = sheet.TabId;

            // If the TabId already exists in the set, a duplicate is found
            if (!tabIdSet.Add(tabId))
            {
                // Throw an exception or handle the duplicate as needed
                throw new InvalidOperationException(
                    $"Duplicate TabId {tabId} detected in worksheet \"{sheet.Name}\".");
            }
        }

        // If no exception was thrown, all TabIds are unique
        // Save the workbook if further processing is required
        workbook.Save("output.xlsx");
    }
}
