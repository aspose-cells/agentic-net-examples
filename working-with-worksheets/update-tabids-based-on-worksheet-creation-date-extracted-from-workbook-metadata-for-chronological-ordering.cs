// Title: Chronologically Reorder Worksheet TabIds Using Revision Metadata – Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, extracts InsertSheet revision timestamps from the workbook's revision logs, assigns a creation date to each worksheet (using DateTime.Max for sheets without metadata), sorts the sheets by these dates, updates their TabId sequentially starting at 0, and saves the modified file.
// Keywords: Aspose.Cells | C# | .NET | Worksheet TabId | reorder worksheets | chronological order | revision logs | InsertSheet revision | SavedTime | Excel workbook | programmatic tab ordering | sheet creation date | update TabId | sample code | GitHub example
// Common Searches: Aspose.Cells set worksheet TabId by creation date | C# sort Excel sheets by insertion time | read revision logs Aspose.Cells | update tab order programmatically .NET | how to reorder worksheet tabs using Aspose.Cells
// Developer Intent: Reorder worksheets based on their original insertion timestamps and assign sequential TabId values using Aspose.Cells in a C# application.
// Use Cases: Restore the original editing sequence of a legacy workbook where sheet creation times are stored in revision logs. | Prepare a workbook for distribution with tabs ordered by creation date for better user navigation. | Merge multiple workbooks and synchronize TabId values to maintain a consistent chronological tab order.
// AI Prompts: Generate C# code with Aspose.Cells that reads InsertSheet revision timestamps, sorts worksheets by those timestamps, and sets TabId sequentially. | Provide a reusable method that accepts a Workbook object and updates each worksheet's TabId based on the earliest revision SavedTime for that sheet. | Explain how to handle worksheets lacking revision metadata when assigning TabId values with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Revisions;

namespace AsposeCellsTabIdUpdater
{
    // Loads an Excel workbook, extracts InsertSheet revision timestamps from the workbook's revision logs, assigns a creation date to each worksheet (using DateTime.Max for sheets without metadata), sorts the sheets by these dates, updates their TabId sequentially starting at 0, and saves the modified file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputPath = "InputWorkbook.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Dictionary to hold creation time for each sheet (key = sheet name)
                var sheetCreationTimes = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

                // RevisionLogs may be null if the workbook has no revisions
                var revisionLogs = workbook.Worksheets.RevisionLogs;
                if (revisionLogs != null)
                {
                    // Iterate through revision logs to find InsertSheet revisions and capture their saved time
                    foreach (RevisionLog log in revisionLogs)
                    {
                        // Metadata may be null; guard against it
                        DateTime revisionTime = log.MetadataTable?.SavedTime ?? DateTime.MinValue;

                        foreach (Revision rev in log.Revisions)
                        {
                            if (rev.Type == RevisionType.InsertSheet && rev is RevisionInsertSheet insertRev)
                            {
                                string sheetName = insertRev.Name;

                                // Store the earliest time we encounter for a given sheet
                                if (!sheetCreationTimes.ContainsKey(sheetName) ||
                                    revisionTime < sheetCreationTimes[sheetName])
                                {
                                    sheetCreationTimes[sheetName] = revisionTime;
                                }
                            }
                        }
                    }
                }

                // Prepare a list of worksheets with their associated creation times
                var worksheetsWithDates = new List<(Worksheet sheet, DateTime created)>();

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // If we couldn't find a creation time, assign a max value so it appears last
                    DateTime created = sheetCreationTimes.TryGetValue(ws.Name, out DateTime dt) ? dt : DateTime.MaxValue;
                    worksheetsWithDates.Add((ws, created));
                }

                // Sort the list chronologically (earliest first)
                var sortedWorksheets = worksheetsWithDates
                    .OrderBy(pair => pair.created)
                    .Select(pair => pair.sheet)
                    .ToList();

                // Update TabId sequentially based on the sorted order
                for (int i = 0; i < sortedWorksheets.Count; i++)
                {
                    sortedWorksheets[i].TabId = i; // Assign TabId starting from 0
                }

                // Output workbook path
                string outputPath = "OutputWorkbook.xlsx";

                // Save the workbook with updated TabIds
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
