using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsGroupingCallbackDemo
{
    // Simple logger for auditing
    public static class AuditLogger
    {
        public static void LogGroupProcessed(string groupId, int recordCount)
        {
            Console.WriteLine($"[Audit] Group ID: {groupId}, Record Count: {recordCount}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (10 rows)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue($"Item {i + 1}");
                cells[i, 1].PutValue(i % 3); // simple category for grouping
            }

            // Define groups (e.g., rows 0‑3, 4‑6, 7‑9)
            var groups = new (int start, int end, string id)[]
            {
                (0, 3, "GroupA"),
                (4, 6, "GroupB"),
                (7, 9, "GroupC")
            };

            // Process each group: group rows and invoke callback
            foreach (var g in groups)
            {
                // Group the rows (hidden = false)
                cells.GroupRows(g.start, g.end, false);

                // Calculate record count for this group
                int recordCount = g.end - g.start + 1;

                // Callback to audit logger
                AuditLogger.LogGroupProcessed(g.id, recordCount);
            }

            // Optional: set outline summary row below groups
            sheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("GroupedRowsWithAudit.xlsx");
        }
    }
}