// Title: How to abort saving an Excel workbook with Aspose.Cells for .NET when any worksheet has a duplicate TabId
// AI Prompts: Write C# code that scans all worksheets in an Aspose.Cells Workbook, logs the name of any sheet with a duplicate TabId, and prevents the workbook.Save() call from executing. | Update the sample to enforce unique TabId values by returning early or throwing an exception if a duplicate is found before attempting to save the file.
// Common Searches: Aspose.Cells .NET check for duplicate worksheet TabId before saving workbook | C# prevent Excel file save when worksheet TabId conflict occurs using Aspose.Cells | validate unique TabId values across worksheets in Aspose.Cells workbook | how to abort workbook.Save() in Aspose.Cells if TabId duplicates are detected
// Tags: Aspose.Cells duplicate TabId detection | C# abort workbook save on TabId conflict | validate worksheet TabId uniqueness Aspose.Cells | Aspose.Cells conditional workbook save | Excel TabId uniqueness check C#

using Aspose.Cells;
using System;
using System.Collections.Generic;

// // Loads an Excel file with Aspose.Cells, optionally modifies worksheets, then iterates through all worksheets using a HashSet to detect duplicate TabId values. If a duplicate is found, it logs the offending sheet and skips the workbook.Save() call; otherwise, it saves the workbook to the specified output file.
class Program
{
    static void Main()
    {
        // Load the workbook (using the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // -------------------------------------------------
        // Apply any required modifications to worksheets here
        // Example: workbook.Worksheets[0].TabId = 10;
        // -------------------------------------------------

        // Detect duplicate TabId values across all worksheets
        HashSet<int> tabIds = new HashSet<int>();
        bool hasDuplicate = false;

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int id = sheet.TabId;
            if (!tabIds.Add(id))
            {
                // Duplicate found
                hasDuplicate = true;
                Console.WriteLine($"Duplicate TabId {id} found on worksheet \"{sheet.Name}\".");
                break;
            }
        }

        // Prevent saving if a duplicate TabId exists
        if (hasDuplicate)
        {
            Console.WriteLine("Workbook will not be saved due to duplicate TabId.");
            return;
        }

        // Save the workbook (using the provided save rule)
        workbook.Save("output.xlsx");
    }
}
