// Title: How to delete rows with duplicate values in columns A and B using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that scans an .xlsx file and eliminates rows where the combination of column A and column B repeats. | Demonstrate using a case‑insensitive HashSet to store A‑B pairs while traversing rows from bottom to top and removing later occurrences. | Create a reusable method that loads a workbook, purges duplicate A‑B entries, and writes the cleaned file to a new location.
// Common Searches: Aspose.Cells C# delete rows when A and B columns have the same values | C# remove repeated records in Excel based on two columns using Aspose | How to filter out duplicate A‑B pairs in an .xlsx file with Aspose.Cells | Programmatic way to clean Excel sheet by dropping rows with identical column A and B values in .NET
// Tags: duplicate A‑B key removal Aspose.Cells | hashset based row deduplication .NET | bottom‑up row deletion Aspose.Cells | excel worksheet data cleaning C# | case‑insensitive key comparison Aspose

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads input.xlsx, iterates from the last row upward, builds a case‑insensitive key from columns A and B, uses a HashSet to detect repeats, deletes any row whose key already exists, and saves the cleaned workbook as output.xlsx.
class RemoveDuplicateRows
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the last row that contains data
        int lastRow = cells.MaxDataRow; // zero‑based index

        // HashSet to store unique combinations of column A and B values
        HashSet<string> uniqueKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Iterate from bottom to top to safely delete rows
        for (int row = lastRow; row >= 0; row--)
        {
            // Read values from column A (index 0) and column B (index 1)
            string valueA = cells[row, 0].StringValue?.Trim() ?? string.Empty;
            string valueB = cells[row, 1].StringValue?.Trim() ?? string.Empty;

            // Combine the two values to form a unique key
            string key = $"{valueA}|{valueB}";

            if (uniqueKeys.Contains(key))
            {
                // Duplicate found – delete the entire row
                worksheet.Cells.DeleteRow(row);
            }
            else
            {
                // First occurrence – add the key to the set
                uniqueKeys.Add(key);
            }
        }

        // Save the cleaned workbook (save rule)
        workbook.Save(outputPath);
    }
}
