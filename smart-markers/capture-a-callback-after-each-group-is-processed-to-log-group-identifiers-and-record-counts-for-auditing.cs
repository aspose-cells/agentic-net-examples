// Title: Log each group identifier and record count while aggregating rows with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells log group identifier after processing rows | how to audit row groups and counts in an Excel workbook using Aspose.Cells | capture callback after each group is processed with Aspose.Cells smart markers | sum column values per group and record audit details in Aspose.Cells | record count per group logging example Aspose.Cells .NET
// Tags: group row aggregation Aspose.Cells | excel audit logging C# | post‑group processing callback Aspose.Cells | write group sum to column C Aspose.Cells | smart markers group handling Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel file, groups rows by the value in column A, calculates the sum of numeric values in column B for each group, writes the sum to column C of the first row in the group, logs the group ID and the number of records processed, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Build groups based on a key column (e.g., column A contains group identifiers)
            var groups = new Dictionary<string, List<int>>();
            int headerRow = 0;               // assuming first row is header
            int startRow = headerRow + 1;    // data starts after header
            int lastRow = sheet.Cells.MaxDataRow;

            for (int row = startRow; row <= lastRow; row++)
            {
                string groupId = sheet.Cells[row, 0].StringValue?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(groupId))
                    continue; // skip rows without a group identifier

                if (!groups.ContainsKey(groupId))
                    groups[groupId] = new List<int>();

                groups[groupId].Add(row);
            }

            // Process each group
            foreach (var kvp in groups)
            {
                string groupId = kvp.Key;
                List<int> rows = kvp.Value;

                // Calculate sum of values in column B for the group
                double sum = 0;
                foreach (int r in rows)
                {
                    Cell cell = sheet.Cells[r, 1]; // column B
                    // Safely retrieve numeric value; treat non‑numeric or empty cells as 0
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        sum += cell.DoubleValue;
                    }
                }

                // Write the sum to column C of the first row of the group
                sheet.Cells[rows[0], 2].PutValue(sum); // column C

                // Log group identifier and record count for auditing
                LogGroupProcessing(groupId, rows.Count);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple audit logger; replace with a proper logging framework if needed
    static void LogGroupProcessing(string groupId, int recordCount)
    {
        Console.WriteLine($"Audit Log - Group ID: {groupId}, Record Count: {recordCount}");
    }
}
