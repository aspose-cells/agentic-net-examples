// Title: C# – Search non‑contiguous ranges G1:G10 and H1:H10 using Aspose.Cells FindOptions.SetRange
// Description: Creates a workbook, populates columns G and H, then demonstrates how to restrict the Find operation to the separate ranges G1:G10 and H1:H10 with FindOptions.SetRange. The example searches for a target string, reports the cell address if found, and saves the file.
// Keywords: Aspose.Cells C# FindOptions | SetRange non contiguous ranges | search specific columns G H | CellArea range search | limit Find to G1:G10 | limit Find to H1:H10 | Aspose.Cells Find example
// Common Searches: Aspose.Cells limit Find to G1:G10 | Search multiple separate ranges C# Aspose.Cells | FindOptions SetRange usage | How to search only columns G and H with Aspose.Cells | Non‑contiguous range search Aspose.Cells
// Developer Intent: Locate a given text exclusively within the two distinct ranges G1:G10 and H1:H10.
// Use Cases: Verify that a required code appears in either column G or H before processing a row. | Highlight cells containing a keyword only in the designated columns without scanning the whole sheet. | Export rows where a term is found in the specified column ranges to another workbook.
// AI Prompts: Generate a method that accepts a list of CellArea objects and a search string, returning all matching cells using FindOptions. | Show how to combine multiple CellArea ranges into a single Find call for non‑contiguous search in Aspose.Cells. | Explain how to retrieve every occurrence of a value across separate ranges such as G1:G10 and H1:H10 in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsSearchNonContiguous
{
    // Creates a workbook, populates columns G and H, then demonstrates how to restrict the Find operation to the separate ranges G1:G10 and H1:H10 with FindOptions.SetRange. The example searches for a target string, reports the cell address if found, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data in columns G and H (indices 6 and 7)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 6].PutValue($"G{i + 1}"); // G1:G10
                cells[i, 7].PutValue($"H{i + 1}"); // H1:H10
            }

            // The value we want to find
            string searchText = "G5";

            // Create FindOptions instance
            FindOptions options = new FindOptions();

            // ---------- First range: G1:G10 ----------
            CellArea rangeG = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                EndRow = 9,        // Row 10
                StartColumn = 6,   // Column G (zero‑based)
                EndColumn = 6      // Single column
            };
            options.SetRange(rangeG);          // Apply the first range
            Cell foundG = cells.Find(searchText, null, options);
            if (foundG != null)
                Console.WriteLine($"Found '{searchText}' at {foundG.Name} in range G1:G10");
            else
                Console.WriteLine($"'{searchText}' not found in range G1:G10");

            // ---------- Second range: H1:H10 ----------
            // Reuse the same FindOptions object; just change the range
            CellArea rangeH = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 7,   // Column H
                EndColumn = 7
            };
            options.SetRange(rangeH);          // Apply the second range
            Cell foundH = cells.Find(searchText, null, options);
            if (foundH != null)
                Console.WriteLine($"Found '{searchText}' at {foundH.Name} in range H1:H10");
            else
                Console.WriteLine($"'{searchText}' not found in range H1:H10");

            // Save the workbook (optional, just to visualize the data)
            workbook.Save("SearchNonContiguous.xlsx");
        }
    }
}
