// Title: Log row group IDs and record counts with a callback after grouping rows – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill it with sample data, define row ranges, group each range using Cells.GroupRows, and invoke a custom LogGroupInfo callback that calculates and logs the start index, end index, and record count for every processed group before saving the file.
// Keywords: Aspose.Cells group rows callback | C# log grouped rows | record count after grouping rows .NET | worksheet row grouping audit | Cells.GroupRows example | Aspose.Cells logging callback | row grouping diagnostics C#
// Common Searches: Aspose.Cells callback after grouping rows | how to log row groups in Aspose.Cells | record count for each grouped row range .NET | audit grouped rows Aspose.Cells C# | custom logging for Cells.GroupRows
// Developer Intent: Capture and log the start/end indices and size of each row group immediately after it is created.
// Use Cases: Maintain an audit trail of grouped sections in financial or reporting spreadsheets. | Validate grouping logic in large data‑processing pipelines by outputting group identifiers and counts. | Debug complex worksheets by printing row‑group details to the console or a log file.
// AI Prompts: Generate a C# method that groups rows with Aspose.Cells and calls a user‑defined callback receiving start and end row indices. | Show how to modify the LogGroupInfo callback to write group information to a text file instead of the console. | Create code that captures grouping events for both rows and columns in Aspose.Cells and logs their ranges and record counts.

using System;
using Aspose.Cells;

namespace AsposeCellsGroupCallbackDemo
{
    // Demonstrates how to create a workbook, fill it with sample data, define row ranges, group each range using Cells.GroupRows, and invoke a custom LogGroupInfo callback that calculates and logs the start index, end index, and record count for every processed group before saving the file.
    class Program
    {
        // Simple callback method that logs group information.
        static void LogGroupInfo(int startIndex, int endIndex)
        {
            int recordCount = endIndex - startIndex + 1;
            Console.WriteLine($"Group processed: Rows {startIndex} to {endIndex} (Count = {recordCount})");
        }

        // Groups rows and invokes the callback after each grouping operation.
        static void GroupRowsWithCallback(Worksheet worksheet, int[][] groups, bool hideGroupedRows)
        {
            Cells cells = worksheet.Cells;

            foreach (int[] range in groups)
            {
                int start = range[0];
                int end = range[1];

                // Perform the actual grouping.
                cells.GroupRows(start, end, hideGroupedRows);

                // Callback logging.
                LogGroupInfo(start, end);
            }
        }

        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (10 rows, 2 columns).
                for (int row = 0; row < 10; row++)
                {
                    cells[row, 0].PutValue($"Item {row + 1}");
                    cells[row, 1].PutValue(row * 10);
                }

                // Define row groups: {0-2}, {3-5}, {6-9}
                int[][] rowGroups = new int[][]
                {
                    new int[] { 0, 2 },
                    new int[] { 3, 5 },
                    new int[] { 6, 9 }
                };

                // Group rows and log after each group.
                GroupRowsWithCallback(sheet, rowGroups, hideGroupedRows: false);

                // Save the workbook.
                string outputPath = "GroupedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
