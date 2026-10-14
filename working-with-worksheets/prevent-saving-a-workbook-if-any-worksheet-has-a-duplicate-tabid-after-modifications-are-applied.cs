// Title: Abort saving an Aspose.Cells workbook in C# when duplicate worksheet TabId values are detected
// AI Prompts: Write C# code using Aspose.Cells that scans every worksheet, identifies duplicate TabId values, logs the offending sheet, and skips the Workbook.Save call if any duplicates are found. | Enhance an existing Aspose.Cells processing pipeline to perform a pre‑save validation that ensures each worksheet's TabId is unique, returning a warning or throwing an exception instead of persisting the file when a conflict occurs.
// Common Searches: c# aspocells prevent workbook.save if two sheets have same tabid | check for duplicate worksheet tabid before saving excel file using Aspose.Cells | how to validate unique tab identifiers in an Aspose.Cells workbook | aspocells detect tabid conflict across worksheets in .NET
// Tags: Aspose.Cells duplicate TabId detection | C# workbook save abort on TabId conflict | validate worksheet TabId uniqueness Aspose | pre‑save TabId validation Aspose.Cells | Excel sheet TabId conflict handling C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace WorkbookTabIdCheck
{
    // The program loads an Excel workbook with Aspose.Cells, optionally modifies worksheets, then iterates through all sheets using a HashSet to ensure each TabId is unique. If a duplicate TabId is found, it logs the conflict and skips the Workbook.Save call, preventing the file from being saved.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // -------------------------------------------------
            // Apply your modifications to the workbook here.
            // For example, you might change worksheet names,
            // add new worksheets, or modify TabId values.
            // -------------------------------------------------
            // Example modification (optional):
            // workbook.Worksheets[0].TabId = 100;

            // Check for duplicate TabId values across all worksheets
            HashSet<int> tabIdSet = new HashSet<int>();
            bool duplicateFound = false;

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int currentTabId = sheet.TabId;

                // If the TabId already exists in the set, we have a duplicate
                if (!tabIdSet.Add(currentTabId))
                {
                    duplicateFound = true;
                    Console.WriteLine($"Duplicate TabId detected: {currentTabId} in worksheet \"{sheet.Name}\".");
                    break; // Stop checking after the first duplicate is found
                }
            }

            // Prevent saving if a duplicate TabId exists
            if (!duplicateFound)
            {
                // Save the workbook (replace with your desired output path)
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            else
            {
                Console.WriteLine("Workbook not saved due to duplicate TabId values.");
            }
        }
    }
}
