// Title: How to enumerate the first row of an Excel worksheet and retrieve distinct column headers with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read first row and get unique header names from Excel | extract distinct column headings from the top row using Aspose.Cells .NET | list unique headers in an Excel sheet with Aspose.Cells C# example | avoid duplicate column names when reading Excel headers with Aspose.Cells | use MaxDataColumn to iterate first‑row cells in Aspose.Cells C#
// Tags: enumerate first row cells Aspose.Cells | retrieve distinct Excel headers C# | case‑insensitive header collection Aspose.Cells | MaxDataColumn iteration Aspose.Cells | unique column names extraction .NET | hashset header deduplication C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, accesses the first worksheet, determines the last populated column via MaxDataColumn, and iterates over each cell in the first row. It trims each cell value, skips blanks, and adds the header to a case‑insensitive HashSet, automatically discarding duplicates, then prints the distinct column headers.
class UniqueColumnHeaders
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;

        try
        {
            // Load the workbook (any loading errors are caught below)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // HashSet to store distinct header values (case‑insensitive)
            HashSet<string> uniqueHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Determine the last column with data in the first row
            int lastColumn = sheet.Cells.MaxDataColumn;

            // Enumerate cells in the first row (row index 0)
            for (int col = 0; col <= lastColumn; col++)
            {
                // Get the cell value as string and trim whitespace
                string header = sheet.Cells[0, col].StringValue?.Trim();

                // Skip empty cells
                if (string.IsNullOrEmpty(header))
                    continue;

                // Add to the set (duplicates are ignored automatically)
                uniqueHeaders.Add(header);
            }

            // Output the unique headers
            Console.WriteLine("Unique column headers:");
            foreach (string header in uniqueHeaders)
            {
                Console.WriteLine(header);
            }

            // Optionally, save the workbook if modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
