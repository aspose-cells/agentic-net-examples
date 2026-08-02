// Title: C# – Log Worksheet TabId Changes with Aspose.Cells
// Description: Demonstrates how to capture the original TabId of each worksheet, assign new TabId values, store the before‑and‑after pairs in a custom log, output the log, save the workbook, and confirm that the changes persist after reloading.
// Keywords: Aspose.Cells TabId logging | C# worksheet TabId audit | track TabId changes .NET | record original TabId Aspose | verify TabId after save | Aspose.Cells workbook modification log
// Common Searches: how to log worksheet TabId changes in Aspose.Cells | record original TabId before modifying C# | Aspose.Cells verify TabId after saving workbook | C# create change log for worksheet TabId | Aspose.Cells track TabId updates
// Developer Intent: Implement a mechanism that records each worksheet's original TabId and its new value whenever the TabId property is modified.
// Use Cases: Maintain an audit trail of TabId modifications during batch worksheet processing. | Generate a console or file report of all TabId updates for compliance or debugging. | Confirm that custom TabId values are retained after the workbook is saved and reopened.
// AI Prompts: Write C# code using Aspose.Cells that logs original and new TabId values for every worksheet when they are changed and prints the log to the console. | Create a method that accepts a Workbook, monitors TabId assignments, stores changes in a list of custom log objects, and returns that list. | Show how to validate that TabId changes are persisted after saving the workbook and reloading it with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTabIdLogger
{
    // Simple logger that records original and new TabId values
    // Demonstrates how to capture the original TabId of each worksheet, assign new TabId values, store the before‑and‑after pairs in a custom log, output the log, save the workbook, and confirm that the changes persist after reloading.
    public class TabIdChangeLog
    {
        public string WorksheetName { get; set; }
        public int OriginalTabId { get; set; }
        public int NewTabId { get; set; }

        public override string ToString()
        {
            return $"Worksheet \"{WorksheetName}\": TabId changed from {OriginalTabId} to {NewTabId}";
        }
    }

    class Program
    {
        static void Main()
        {
            // List to hold change logs
            List<TabIdChangeLog> changeLogs = new List<TabIdChangeLog>();

            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook(); // using Aspose.Cells Workbook constructor
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";

            // Add a second worksheet for demonstration
            int secondIndex = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[secondIndex];
            sheet2.Name = "SecondSheet";

            // ---------- Record original TabId values ----------
            int originalTabId1 = sheet1.TabId;
            int originalTabId2 = sheet2.TabId;

            // ---------- Change TabId values ----------
            sheet1.TabId = 101; // new value
            sheet2.TabId = 202; // new value

            // ---------- Log the changes ----------
            changeLogs.Add(new TabIdChangeLog
            {
                WorksheetName = sheet1.Name,
                OriginalTabId = originalTabId1,
                NewTabId = sheet1.TabId
            });

            changeLogs.Add(new TabIdChangeLog
            {
                WorksheetName = sheet2.Name,
                OriginalTabId = originalTabId2,
                NewTabId = sheet2.TabId
            });

            // ---------- Output the log ----------
            Console.WriteLine("TabId Change Log:");
            foreach (var log in changeLogs)
            {
                Console.WriteLine(log);
            }

            // ---------- Save the workbook ----------
            string outputPath = "TabIdDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx); // using Aspose.Cells Save method

            // ---------- Load the saved workbook and verify TabId ----------
            Workbook loadedWorkbook = new Workbook(outputPath); // using Aspose.Cells load constructor
            Worksheet loadedSheet1 = loadedWorkbook.Worksheets[0];
            Worksheet loadedSheet2 = loadedWorkbook.Worksheets[1];

            Console.WriteLine("\nVerification after loading:");
            Console.WriteLine($"Loaded Worksheet \"{loadedSheet1.Name}\" TabId: {loadedSheet1.TabId}");
            Console.WriteLine($"Loaded Worksheet \"{loadedSheet2.Name}\" TabId: {loadedSheet2.TabId}");
        }
    }
}
