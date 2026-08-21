// Title: Aspose.Cells .NET: Log Worksheet TabId Changes with Timestamp
// Description: This example introduces a TabIdChangeLogger class that records the worksheet name, previous TabId, new TabId, and the exact time of each change. A helper method updates a worksheet's TabId only when the value differs and automatically adds an entry to the logger. The sample shows logging before and after saving a workbook and prints a formatted change report to the console.
// Keywords: Aspose.Cells | .NET | C# | Worksheet TabId | log TabId changes | audit workbook tabs | timestamped change log | track tab order | Excel workbook modification history | TabIdChangeLogger
// Common Searches: Aspose.Cells log TabId changes | C# track worksheet TabId modifications | record old and new TabId values Aspose | audit Excel tab order with Aspose.Cells | how to log worksheet identifier changes .NET
// Developer Intent: Automatically capture every alteration of a worksheet's TabId, storing the old value, the new value, the worksheet name, and a timestamp for auditing or debugging purposes.
// Use Cases: Generate an audit trail of TabId updates when programmatically reordering worksheet tabs. | Detect unintended TabId changes after loading a saved workbook before further processing. | Produce a compliance report that lists all TabId modifications across multiple worksheets. | Integrate TabId change data with existing logging frameworks or monitoring dashboards.
// AI Prompts: Create code to export TabIdChangeLogger entries to a CSV file with columns for worksheet name, old TabId, new TabId, and timestamp. | Show how to wrap the Worksheet class so that any assignment to TabId automatically triggers the logger without explicit SetTabId calls. | Write unit tests for TabIdChangeLogger that verify entries are added only when the TabId actually changes. | Demonstrate sending TabId change events to a centralized logging service such as Serilog or Azure Application Insights.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTabIdLogger
{
    // Simple logger that records original and new TabId values
    // This example introduces a TabIdChangeLogger class that records the worksheet name, previous TabId, new TabId, and the exact time of each change. A helper method updates a worksheet's TabId only when the value differs and automatically adds an entry to the logger. The sample shows logging before and after saving a workbook and prints a formatted change report to the console.
    public class TabIdChangeLogger
    {
        // Each log entry stores worksheet name, old TabId and new TabId
        public class LogEntry
        {
            public string WorksheetName { get; set; }
            public int OldTabId { get; set; }
            public int NewTabId { get; set; }
            public DateTime Timestamp { get; set; }
        }

        private readonly List<LogEntry> _entries = new List<LogEntry>();

        // Record a change
        public void Record(Worksheet sheet, int oldId, int newId)
        {
            _entries.Add(new LogEntry
            {
                WorksheetName = sheet.Name,
                OldTabId = oldId,
                NewTabId = newId,
                Timestamp = DateTime.Now
            });
        }

        // Output all logged changes to console
        public void PrintLog()
        {
            Console.WriteLine("=== TabId Change Log ===");
            foreach (var e in _entries)
            {
                Console.WriteLine($"{e.Timestamp:u} - Worksheet \"{e.WorksheetName}\": TabId changed from {e.OldTabId} to {e.NewTabId}");
            }
            Console.WriteLine("========================");
        }
    }

    class Program
    {
        // Helper method that changes TabId and logs the change
        static void SetTabId(Worksheet sheet, int newTabId, TabIdChangeLogger logger)
        {
            int oldTabId = sheet.TabId;
            if (oldTabId != newTabId)
            {
                sheet.TabId = newTabId;
                logger.Record(sheet, oldTabId, newTabId);
            }
        }

        static void Main()
        {
            // Initialize logger
            var logger = new TabIdChangeLogger();

            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook(); // create rule
            Worksheet ws = workbook.Worksheets[0];

            // Change TabId a few times, logging each change
            SetTabId(ws, 101, logger);
            SetTabId(ws, 202, logger);
            SetTabId(ws, 202, logger); // no change, won't be logged
            SetTabId(ws, 303, logger);

            // Save the workbook
            string filePath = "TabIdDemo.xlsx";
            workbook.Save(filePath); // save rule

            // ---------- Load the saved workbook ----------
            Workbook loadedWorkbook = new Workbook(filePath); // load rule
            Worksheet loadedWs = loadedWorkbook.Worksheets[0];

            // Change TabId after loading, logging the change
            SetTabId(loadedWs, 404, logger);
            SetTabId(loadedWs, 505, logger);

            // Save again
            string updatedPath = "TabIdDemo_Updated.xlsx";
            loadedWorkbook.Save(updatedPath); // save rule

            // Print the log of all TabId changes
            logger.PrintLog();
        }
    }
}
