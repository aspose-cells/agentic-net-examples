using System;
using Aspose.Cells;

namespace AsposeCellsSearchNonContiguous
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in the two non‑contiguous ranges G1:G10 and H1:H10
            for (int i = 0; i < 10; i++)
            {
                // Column G (index 6)
                cells[i, 6].PutValue($"G{i + 1}");
                // Column H (index 7)
                cells[i, 7].PutValue($"H{i + 1}");
            }

            // Add a few extra values to demonstrate the search limit
            cells["A1"].PutValue("G5"); // Same value but outside the target ranges
            cells["F5"].PutValue("H7"); // Same value but outside the target ranges

            // Configure FindOptions to limit the search to G1:G10 and H1:H10
            FindOptions options = new FindOptions();

            // Define a CellArea that covers both columns G and H from row 1 to row 10
            // Row and column indexes are zero‑based.
            CellArea searchArea = new CellArea
            {
                StartRow = 0,      // Row 1
                EndRow = 9,        // Row 10
                StartColumn = 6,   // Column G
                EndColumn = 7      // Column H
            };

            // Apply the range to the FindOptions
            options.SetRange(searchArea);

            // Example: find the first occurrence of the text "G5" within the defined area
            Cell found = sheet.Cells.Find("G5", null, options);

            if (found != null)
            {
                Console.WriteLine($"Found 'G5' at {found.Name} (Row {found.Row + 1}, Column {found.Column + 1})");
            }
            else
            {
                Console.WriteLine("Value 'G5' not found within the specified non‑contiguous ranges.");
            }

            // Save the workbook (optional, just to visualize the data)
            workbook.Save("SearchNonContiguousRanges.xlsx");
        }
    }
}