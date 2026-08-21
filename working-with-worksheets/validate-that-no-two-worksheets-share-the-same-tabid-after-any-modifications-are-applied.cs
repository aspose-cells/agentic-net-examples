// Title: Validate and Resolve Duplicate Worksheet TabId Values in Aspose.Cells (C#)
// Description: This example creates a workbook with three sheets, intentionally assigns the same TabId to two of them, and then runs the EnsureUniqueTabIds method. The routine scans every worksheet, records used TabIds with a HashSet, finds the highest existing TabId, and assigns new incremental identifiers to any duplicates before saving the file.
// Keywords: Aspose.Cells C# example | worksheet TabId validation | duplicate TabId detection | unique sheet tab identifier | HashSet usage in Aspose.Cells | max TabId calculation | Excel tab collision fix | GitHub source code | global development | US developers | European C# community
// Common Searches: how to check for duplicate TabId in Aspose.Cells | C# code to make worksheet TabIds unique | Aspose.Cells resolve sheet tab ID conflict | prevent Excel tab duplication with Aspose | sample project for TabId validation on GitHub
// Developer Intent: Ensure that each worksheet in a workbook has a distinct TabId.
// Use Cases: After programmatically adding or renaming sheets, invoke EnsureUniqueTabIds to avoid TabId clashes before exporting the workbook. | When merging worksheets from external files, validate TabIds to preserve correct tab order and prevent UI anomalies in Excel. | Prior to distributing a workbook to end‑users, enforce unique TabIds to eliminate rendering issues caused by duplicate sheet tabs.
// AI Prompts: Generate a C# method for Aspose.Cells that scans all worksheets and reassigns new TabId values to any duplicates while keeping existing unique IDs unchanged. | Write code that logs the names of worksheets with conflicting TabIds in an Aspose.Cells workbook and then fixes the conflicts. | Refactor the EnsureUniqueTabIds routine to use LINQ for detecting duplicate TabIds and assign sequential IDs starting from the current maximum.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTabIdValidation
{
    // This example creates a workbook with three sheets, intentionally assigns the same TabId to two of them, and then runs the EnsureUniqueTabIds method. The routine scans every worksheet, records used TabIds with a HashSet, finds the highest existing TabId, and assigns new incremental identifiers to any duplicates before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Add a third worksheet
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Manually set TabId values to create a duplicate scenario
            sheet1.TabId = 100;
            sheet2.TabId = 200;
            sheet3.TabId = 100; // Duplicate TabId with sheet1

            // Validate and resolve duplicate TabIds
            EnsureUniqueTabIds(workbook);

            // Save the workbook (adjust path as needed)
            workbook.Save("ValidatedWorkbook.xlsx");
        }

        /// <param name="workbook">The workbook to validate.</param>
        static void EnsureUniqueTabIds(Workbook workbook)
        {
            // Keep track of used TabIds
            HashSet<int> usedTabIds = new HashSet<int>();

            // Determine the maximum existing TabId to start generating new ones
            int maxTabId = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.TabId > maxTabId)
                    maxTabId = ws.TabId;
            }

            // Iterate through worksheets and resolve duplicates
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (usedTabIds.Contains(ws.TabId))
                {
                    // Duplicate found – assign a new unique TabId
                    maxTabId++;
                    ws.TabId = maxTabId;
                }

                // Record the (now unique) TabId
                usedTabIds.Add(ws.TabId);
            }
        }
    }
}
