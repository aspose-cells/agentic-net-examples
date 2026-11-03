// Title: How to remove duplicate rows from a CSV file using a key column with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# delete duplicate rows in CSV based on first column | remove rows with duplicate keys from CSV using Aspose.Cells workbook | C# Aspose.Cells example for CSV deduplication by column index | filter duplicate records in CSV file with Aspose.Cells and HashSet | save cleaned CSV after removing duplicate rows using Aspose.Cells .NET
// Tags: CSV duplicate row removal Aspose.Cells | key column deduplication .NET | Aspose.Cells delete rows CSV | hashset duplicate detection C# | save workbook as CSV Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells Workbook, scans each row using a specified key column index, tracks encountered keys with a case‑insensitive HashSet, marks rows with duplicate keys for deletion, removes those rows from the worksheet, and saves the cleaned data to a new CSV file.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string inputPath = "input.csv";

        // Path to the destination CSV file after removing duplicates
        string outputPath = "output.csv";

        // Index of the key column used to detect duplicates (0‑based)
        int keyColumnIndex = 0;

        // Load the CSV file into a workbook
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (CSV data is loaded into the first sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Get the used range of the worksheet
        Cells cells = sheet.Cells;
        int maxRow = cells.MaxDataRow; // zero‑based index of the last row with data

        // HashSet to store seen key values
        HashSet<string> seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // List to hold row indices that need to be removed
        List<int> rowsToDelete = new List<int>();

        // Iterate through rows to identify duplicates
        for (int row = 0; row <= maxRow; row++)
        {
            // Retrieve the key cell value as string
            string key = cells[row, keyColumnIndex]?.StringValue?.Trim() ?? string.Empty;

            // If the key is empty, treat it as a unique row (or adjust logic as needed)
            if (string.IsNullOrEmpty(key))
                continue;

            // If the key has been seen before, mark the row for deletion
            if (seenKeys.Contains(key))
            {
                rowsToDelete.Add(row);
            }
            else
            {
                // First occurrence of this key – add to the set
                seenKeys.Add(key);
            }
        }

        // Delete rows starting from the bottom to preserve indices
        for (int i = rowsToDelete.Count - 1; i >= 0; i--)
        {
            int rowIndex = rowsToDelete[i];
            cells.DeleteRow(rowIndex);
        }

        // Save the cleaned data back to a CSV file
        workbook.Save(outputPath, SaveFormat.CSV);
    }
}
