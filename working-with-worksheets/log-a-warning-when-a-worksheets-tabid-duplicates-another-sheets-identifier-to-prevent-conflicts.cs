// Title: Detect and Log Duplicate Worksheet TabId Values with Aspose.Cells for .NET
// Description: Creates a workbook, assigns explicit TabId values to three worksheets (including a duplicate), scans all sheets, logs a console warning for any TabId that appears on multiple worksheets, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | Worksheet TabId | duplicate TabId detection | log warning | workbook validation | sheet identifier conflict | Aspose.Cells example
// Common Searches: Aspose.Cells duplicate TabId warning | how to check for duplicate worksheet TabId in C# | detect repeated TabId in Aspose.Cells workbook | log worksheet identifier conflicts Aspose.Cells | validate unique TabId across sheets .NET
// Developer Intent: Identify and report worksheets that share the same TabId to prevent identifier conflicts.
// Use Cases: Run a pre‑publish check that ensures every worksheet has a unique TabId and outputs warnings for any duplicates. | Generate a diagnostic list of sheets with colliding TabIds to troubleshoot navigation or macro issues. | Integrate duplicate TabId detection into automated build pipelines to avoid runtime errors in applications relying on unique identifiers.
// AI Prompts: Create a reusable C# method that receives a Workbook and returns groups of worksheet names that share the same TabId. | Show how to replace the console warning with a custom exception that includes the duplicate TabId and the affected worksheet names. | Demonstrate logging duplicate TabId warnings to a file using Aspose.Cells together with Microsoft.Extensions.Logging.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Creates a workbook, assigns explicit TabId values to three worksheets (including a duplicate), scans all sheets, logs a console warning for any TabId that appears on multiple worksheets, and saves the file.
public class DuplicateTabIdWarningDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the default first worksheet and set its TabId
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "Sheet1";
        ws1.TabId = 101;

        // Add a second worksheet with a unique TabId
        Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
        ws2.TabId = 102;

        // Add a third worksheet that intentionally duplicates the TabId of the first sheet
        Worksheet ws3 = workbook.Worksheets.Add("Sheet3");
        ws3.TabId = 101; // Duplicate TabId

        // Detect duplicate TabId values across all worksheets
        var tabIdMap = new Dictionary<int, List<string>>();
        foreach (Worksheet ws in workbook.Worksheets)
        {
            int id = ws.TabId;
            if (!tabIdMap.ContainsKey(id))
            {
                tabIdMap[id] = new List<string>();
            }
            tabIdMap[id].Add(ws.Name);
        }

        // Log a warning for each duplicate TabId found
        foreach (var entry in tabIdMap)
        {
            if (entry.Value.Count > 1)
            {
                Console.WriteLine($"Warning: Duplicate TabId {entry.Key} found in worksheets: {string.Join(", ", entry.Value)}");
            }
        }

        // Save the workbook
        string outputPath = "DuplicateTabIdDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
