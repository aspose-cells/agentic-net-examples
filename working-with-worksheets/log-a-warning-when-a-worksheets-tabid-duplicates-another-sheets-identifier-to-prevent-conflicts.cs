// Title: Log a warning for duplicate worksheet TabId values in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# routine that uses Aspose.Cells to iterate all worksheets and output a console warning whenever two sheets share the same TabId. | Provide a reusable function that returns a collection of worksheet names that have non‑unique TabId values in a given workbook.
// Common Searches: asp.net find worksheets with same TabId using Aspose.Cells | c# code to warn about duplicate sheet identifiers in Excel | how to check for TabId conflicts in an Excel workbook with Aspose.Cells | detect duplicate sheet TabId values in .NET
// Tags: detect worksheet TabId collisions Aspose.Cells | log console warning for TabId duplicates .NET | validate Excel sheet identifiers using Aspose.Cells | ensure unique TabId values in workbook

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, tracks TabId values in a dictionary, and writes a console warning when a duplicate TabId is encountered, showing the names of the conflicting worksheets.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Keep track of TabId values and the sheet that first used each one
        Dictionary<int, string> tabIdMap = new Dictionary<int, string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int tabId = sheet.TabId;

            if (tabIdMap.ContainsKey(tabId))
            {
                // Duplicate TabId detected – log a warning
                Console.WriteLine(
                    $"Warning: Worksheet '{sheet.Name}' has duplicate TabId {tabId} (already used by '{tabIdMap[tabId]}').");
            }
            else
            {
                // Record the TabId for future duplicate checks
                tabIdMap[tabId] = sheet.Name;
            }
        }

        // Optionally save the workbook if modifications were made
        // workbook.Save("output.xlsx");
    }
}
