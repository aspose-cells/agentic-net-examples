// Title: Detect and list duplicate text entries in a specific Excel column using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates a chosen column, and returns a dictionary of text values with their occurrence counts, highlighting those occurring more than once. | Write a method using Aspose.Cells that trims cell strings, performs case‑insensitive counting, and prints each duplicate value together with its frequency. | Create a reusable function in C# that accepts a worksheet and column index, enumerates the cells with Aspose.Cells, and outputs duplicate entries while ignoring empty cells.
// Common Searches: Aspose.Cells C# find duplicate strings in column A of an Excel file | C# enumerate Excel column and count repeated text values with Aspose.Cells | case insensitive duplicate detection in Excel using Aspose.Cells .NET | how to get last data row and check for duplicate entries in a column with Aspose.Cells
// Tags: duplicate text detection Aspose.Cells .NET | enumerate column cells Aspose.Cells | case‑insensitive string counting Excel C# | max data row retrieval Aspose.Cells | dictionary occurrence tracking Excel column

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through all rows of a specified column, trims and case‑insensitively counts each non‑empty string using a Dictionary, and prints any values that appear more than once.
class DetectDuplicateTextInColumn
{
    static void Main(string[] args)
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Load the workbook (using the provided load rule)
        Workbook workbook = new Workbook(filePath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Specify the column index to check (0 = column A)
        int columnIndex = 0;

        // Dictionary to track occurrences of each text value
        Dictionary<string, int> occurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Determine the last row with data in the worksheet
        int lastRow = sheet.Cells.MaxDataRow;

        // Enumerate the cells in the specified column
        for (int row = 0; row <= lastRow; row++)
        {
            // Get the cell value as a string
            string cellValue = sheet.Cells[row, columnIndex].StringValue?.Trim();

            // Skip empty cells
            if (string.IsNullOrEmpty(cellValue))
                continue;

            // Update the count for this text entry
            if (occurrences.ContainsKey(cellValue))
                occurrences[cellValue]++;
            else
                occurrences[cellValue] = 1;
        }

        // Output duplicate entries
        Console.WriteLine("Duplicate text entries in column {0}:", (char)('A' + columnIndex));
        bool anyDuplicates = false;
        foreach (var kvp in occurrences)
        {
            if (kvp.Value > 1)
            {
                anyDuplicates = true;
                Console.WriteLine($"\"{kvp.Key}\" occurs {kvp.Value} times");
            }
        }

        if (!anyDuplicates)
        {
            Console.WriteLine("No duplicate text entries found.");
        }

        // (Optional) Save the workbook if modifications were made
        // workbook.Save("output.xlsx"); // using the provided save rule
    }
}
