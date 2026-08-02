// Title: Log a Warning for Duplicate Worksheet TabId in Aspose.Cells (.NET)
// Description: Creates a workbook with three worksheets, assigns TabId values, scans all sheets with a dictionary, writes a console warning when a TabId is reused, and saves the file.
// Keywords: Aspose.Cells duplicate TabId | C# worksheet TabId conflict | detect duplicate sheet identifier | log warning Aspose.Cells | TabId uniqueness check
// Common Searches: Aspose.Cells find duplicate TabId | C# log warning for repeated worksheet TabId | prevent sheet TabId clash in .NET | how to detect duplicate TabId in Excel workbook | validate worksheet identifiers before saving
// Developer Intent: Identify worksheets that share the same TabId and output a warning to avoid identifier collisions.
// Use Cases: Validate TabId uniqueness during automated report generation. | Detect and log duplicate TabId values in batch workbook processing. | Ensure UI tab navigation works correctly by preventing TabId conflicts before publishing.
// AI Prompts: Generate a reusable C# method that returns all duplicate TabId pairs in a Workbook. | Rewrite the sample to throw an exception instead of logging when a duplicate TabId is found. | Create a utility class for Aspose.Cells that checks TabId uniqueness and logs detailed messages for each conflict.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Creates a workbook with three worksheets, assigns TabId values, scans all sheets with a dictionary, writes a console warning when a TabId is reused, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the default first worksheet and set a TabId
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "Sheet1";
        ws1.TabId = 101;

        // Add a second worksheet with a different TabId
        Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
        ws2.TabId = 102;

        // Add a third worksheet that intentionally uses a duplicate TabId
        Worksheet ws3 = workbook.Worksheets.Add("Sheet3");
        ws3.TabId = 101; // Duplicate TabId

        // Detect duplicate TabId values across all worksheets
        var tabIdLookup = new Dictionary<int, string>();
        foreach (Worksheet ws in workbook.Worksheets)
        {
            int currentId = ws.TabId;
            if (tabIdLookup.ContainsKey(currentId))
            {
                // Log a warning when a duplicate TabId is found
                Console.WriteLine($"Warning: Worksheet \"{ws.Name}\" has duplicate TabId {currentId} (already used by \"{tabIdLookup[currentId]}\").");
            }
            else
            {
                tabIdLookup[currentId] = ws.Name;
            }
        }

        // Save the workbook
        workbook.Save("DuplicateTabIdDemo.xlsx");
    }
}
