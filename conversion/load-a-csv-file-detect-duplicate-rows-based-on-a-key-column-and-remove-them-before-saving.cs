// Title: How to remove duplicate rows from a CSV file using a key column with Aspose.Cells for .NET
// Common Searches: C# code to delete duplicate rows in a CSV based on the first column using Aspose.Cells | Aspose.Cells .NET remove duplicate records from CSV file by key column | Load CSV with Aspose.Cells, deduplicate rows, and save as new CSV in C# | Using HashSet to filter duplicate rows when processing CSV with Aspose.Cells | Aspose.Cells example for deleting rows while iterating from bottom to top
// Tags: duplicate row removal CSV Aspose.Cells | key column deduplication .NET | load CSV LoadOptions Aspose.Cells | delete rows programmatically Aspose.Cells | save CSV SaveFormat.Csv Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a CSV file with Aspose.Cells, iterates rows from bottom to top, uses a case‑insensitive HashSet to track values from a specified key column, deletes rows whose key has already been seen, ensures the output directory exists, and saves the cleaned data as a new CSV file.
class Program
{
    static void Main()
    {
        // Paths for input and output CSV files
        string inputPath = "input.csv";
        string outputPath = "output.csv";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file with appropriate load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // CSV is loaded into the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Zero‑based index of the column used as the duplicate key (e.g., 0 = column A)
            int keyColumnIndex = 0;

            // Track keys that have already been encountered
            HashSet<string> seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Determine the total number of used rows (MaxDataRow is zero‑based)
            int totalRows = cells.MaxDataRow + 1;

            // Iterate from bottom to top so row deletions do not affect the loop index
            for (int row = totalRows - 1; row >= 0; row--)
            {
                // Retrieve the key value from the designated column
                string key = cells[row, keyColumnIndex].StringValue?.Trim() ?? string.Empty;

                // Skip empty keys (treat as unique)
                if (string.IsNullOrEmpty(key))
                    continue;

                if (seenKeys.Contains(key))
                {
                    // Delete the duplicate row
                    cells.DeleteRows(row, 1);
                }
                else
                {
                    // First occurrence – add to the set
                    seenKeys.Add(key);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the cleaned data back to CSV format
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Duplicate removal completed. Output saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
