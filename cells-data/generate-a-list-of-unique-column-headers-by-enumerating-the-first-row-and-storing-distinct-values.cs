// Title: C# – Extract Unique Column Headers from First Row Using Aspose.Cells
// Description: Demonstrates how to create a workbook, write a header row with duplicates, then iterate the first row up to the last populated column, trim values, skip blanks, and collect case‑insensitive distinct headers while preserving their original order. The unique headers are printed and the workbook can be saved.
// Keywords: Aspose.Cells C# | unique column headers | distinct worksheet headers | first row enumeration | case‑insensitive header collection | remove duplicate columns | Aspose.Cells .NET example
// Common Searches: Aspose.Cells get unique headers C# | C# extract distinct column names first row | remove duplicate worksheet headers Aspose.Cells | enumerate first row cells Aspose.Cells .NET | case insensitive header list Aspose.Cells
// Developer Intent: Retrieve distinct header names from the first worksheet row, maintaining the order they appear.
// Use Cases: Validate template files by ensuring column names are not duplicated before data import. | Create a mapping of column indexes to unique header strings for dynamic data processing. | Generate a summary report or UI dropdown that lists only the unique column titles. | Export the unique header list to another workbook or external system.
// AI Prompts: Write C# code with Aspose.Cells that reads the first row of a worksheet and returns a List<string> of unique headers, ignoring case and preserving order. | Show how to log the extracted unique headers to the console and optionally save the workbook. | Explain how to adapt the example to handle merged header cells or multi‑row headers while still producing a distinct list.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to create a workbook, write a header row with duplicates, then iterate the first row up to the last populated column, trim values, skip blanks, and collect case‑insensitive distinct headers while preserving their original order. The unique headers are printed and the workbook can be saved.
class UniqueHeadersDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample header row with duplicate values
        string[] sampleHeaders = { "Name", "Age", "Name", "Email", "Age" };
        for (int col = 0; col < sampleHeaders.Length; col++)
        {
            cells[0, col].PutValue(sampleHeaders[col]);
        }

        // Collect distinct header values from the first row
        HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        List<string> uniqueHeaders = new List<string>();

        int lastColumn = cells.MaxDataColumn; // last column with data in the sheet
        for (int col = 0; col <= lastColumn; col++)
        {
            string header = cells[0, col].StringValue?.Trim();
            if (string.IsNullOrEmpty(header))
                continue; // skip empty cells

            // Add to set; if added successfully, also add to list to preserve order
            if (seen.Add(header))
            {
                uniqueHeaders.Add(header);
            }
        }

        // Display the unique headers
        Console.WriteLine("Unique column headers:");
        foreach (string header in uniqueHeaders)
        {
            Console.WriteLine(header);
        }

        // Save the workbook (optional)
        workbook.Save("UniqueHeadersDemo.xlsx");
    }
}
