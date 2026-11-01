// Title: Log a warning for duplicate worksheet TabId values in an Aspose.Cells workbook (C#)
// AI Prompts: Generate C# code with Aspose.Cells that iterates every worksheet, identifies repeated TabId values, and writes a console warning showing the conflicting sheet names. | Create a reusable C# method that accepts a Workbook object, checks for duplicate TabId identifiers across all worksheets, and logs each conflict.
// Common Searches: asp.net aspose.cells detect duplicate worksheet tabid c# | c# code to log warning when two sheets share the same TabId in Aspose.Cells | how to find repeated TabId values across worksheets using Aspose.Cells | check for worksheet TabId collisions in an Excel file with Aspose.Cells C#
// Tags: detect duplicate worksheet TabId Aspose.Cells | log worksheet TabId conflicts C# | validate unique sheet TabId Aspose.Cells | iterate worksheets TabId map C# | Aspose.Cells duplicate TabId warning

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, builds a dictionary of TabId to worksheet name, scans each worksheet, and writes a console warning whenever a TabId is already used by another sheet, then saves the workbook.
class WorksheetTabIdChecker
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to keep track of TabId values and the first worksheet that uses each TabId
        Dictionary<int, string> tabIdMap = new Dictionary<int, string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int tabId = sheet.TabId;

            // Check if this TabId has already been encountered
            if (tabIdMap.ContainsKey(tabId))
            {
                // Log a warning about the duplicate TabId
                Console.WriteLine(
                    $"Warning: Worksheet '{sheet.Name}' has duplicate TabId {tabId} (already used by '{tabIdMap[tabId]}').");
            }
            else
            {
                // Record the TabId and associated worksheet name
                tabIdMap[tabId] = sheet.Name;
            }
        }

        // Save the workbook if any changes were made (optional)
        workbook.Save("output.xlsx");
    }
}
