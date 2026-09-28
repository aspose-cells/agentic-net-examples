// Title: Log original and new Worksheet TabId values when updating them with Aspose.Cells for .NET
// AI Prompts: Create a C# logger class that records the previous TabId and the new TabId each time the Worksheet.TabId property is set in Aspose.Cells. | Add code to export the accumulated TabId change entries to a text file after processing a workbook.
// Common Searches: c# Aspose.Cells how to track worksheet TabId changes and generate a log file | record old and new TabId for each sheet when modifying a workbook with Aspose.Cells | save worksheet TabId modification history to a text file using Aspose.Cells .NET
// Tags: Aspose.Cells worksheet TabId audit | C# logger for worksheet TabId changes | export TabId change history to file | programmatic TabId assignment Aspose.Cells | record worksheet identifier updates .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example defines a WorksheetTabIdLogger that captures the original and new TabId for each worksheet whenever the TabId property is changed, stores the messages in a list, and can write the log to a text file. The main program loads a workbook, updates each sheet's TabId, saves the workbook, and persists the change log.
class WorksheetTabIdLogger
{
    // Stores log entries
    private readonly List<string> _log = new List<string>();

    // Sets a new TabId and records the change
    public void SetTabId(Worksheet ws, int newTabId)
    {
        int oldTabId = ws.TabId;
        if (oldTabId != newTabId)
        {
            ws.TabId = newTabId;
            _log.Add($"Worksheet '{ws.Name}': TabId changed from {oldTabId} to {newTabId}.");
        }
    }

    // Writes the log to a file
    public void SaveLog(string filePath)
    {
        File.WriteAllLines(filePath, _log);
    }

    // Returns the log entries (optional)
    public IEnumerable<string> GetLog()
    {
        return _log;
    }
}

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Initialize the logger
        WorksheetTabIdLogger logger = new WorksheetTabIdLogger();

        // Example: change TabId for each worksheet and log the change
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // New TabId logic (here we just add 100 to the worksheet index)
            int newTabId = ws.Index + 100;
            logger.SetTabId(ws, newTabId);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");

        // Persist the log to a text file
        logger.SaveLog("TabIdChanges.log");
    }
}
