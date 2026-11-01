// Title: How to log group identifiers and record counts with a callback while iterating Excel rows using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells iterate rows and log each group ID and record count | Aspose.Cells callback after processing a group of rows for audit logging | Capture group boundaries in an Excel sheet with Aspose.Cells and output counts | Implement auditing of grouped rows in Excel using Aspose.Cells smart markers C# | Log group identifier and row count during Excel processing with Aspose.Cells .NET
// Tags: Aspose.Cells group row iteration callback | C# Excel group processing audit | Aspose.Cells log group identifier | smart markers group count logging | Aspose.Cells callback for row groups

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook if missing, loads it with Aspose.Cells, scans column A to detect changes in group identifiers, and invokes a LogGroup callback that writes each group's ID and the number of rows processed. The workbook is then saved, providing a reusable pattern for auditing grouped data in Excel using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                var cells = ws.Cells;

                // Sample data: column A contains group identifiers.
                cells[0, 0].PutValue("Group1");
                cells[1, 0].PutValue("Group1");
                cells[2, 0].PutValue("Group2");
                cells[3, 0].PutValue("Group2");
                cells[4, 0].PutValue("Group2");
                cells[5, 0].PutValue("Group3");

                wb.Save(inputPath);
            }

            // Load the workbook (lifecycle rule)
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];
            var cellsRef = worksheet.Cells;

            // Assume rows are grouped by the value in column A (index 0)
            int startRow = 0;                         // first data row (zero‑based)
            int lastRow = cellsRef.MaxDataRow;        // last row that contains data

            string currentGroupId = null;             // identifier of the current group
            int currentGroupCount = 0;                // number of records in the current group

            // Iterate through rows to detect group boundaries
            for (int row = startRow; row <= lastRow; row++)
            {
                // Read the group identifier from column A
                string groupId = cellsRef[row, 0].StringValue?.Trim();

                // Skip rows without a group identifier
                if (string.IsNullOrEmpty(groupId))
                    continue;

                if (currentGroupId == null)
                {
                    // First group encountered
                    currentGroupId = groupId;
                    currentGroupCount = 1;
                }
                else if (groupId == currentGroupId)
                {
                    // Still within the same group
                    currentGroupCount++;
                }
                else
                {
                    // Group changed – invoke the callback for the completed group
                    LogGroup(currentGroupId, currentGroupCount);

                    // Start tracking the new group
                    currentGroupId = groupId;
                    currentGroupCount = 1;
                }
            }

            // After the loop, log the final group if any
            if (currentGroupId != null)
            {
                LogGroup(currentGroupId, currentGroupCount);
            }

            // Save the workbook (lifecycle rule)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log unexpected errors without crashing the application
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Callback that logs the group identifier and the number of records processed
    static void LogGroup(string groupId, int recordCount)
    {
        // Replace this with actual auditing logic (e.g., write to a file or database)
        Console.WriteLine($"Group '{groupId}' processed with {recordCount} records.");
    }
}
