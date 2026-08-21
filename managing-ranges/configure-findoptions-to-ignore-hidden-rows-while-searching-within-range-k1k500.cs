// Title: Aspose.Cells C# FindOptions – Search K1:K500 while ignoring hidden rows
// Description: C# example that loads a workbook, sets a FindOptions range to K1:K500, and iteratively searches for a target string. Each match is checked with Cells.IsRowHidden so only visible rows are considered. The code reports the first visible occurrence or indicates that none was found.
// Keywords: Aspose.Cells FindOptions ignore hidden rows | C# search column K range K1:K500 | Aspose.Cells skip hidden rows | FindOptions SetRange Aspose.Cells .NET | Cells.IsRowHidden example | search visible cells Aspose.Cells | Aspose.Cells find value in specific range
// Common Searches: Aspose.Cells find value in column K ignoring hidden rows | C# FindOptions SetRange K1:K500 | How to skip hidden rows with Aspose.Cells Find | Search visible cells only Aspose.Cells .NET | Find first visible occurrence in a range using Aspose.Cells
// Developer Intent: Locate the first visible occurrence of a given text in column K (rows 1‑500) without considering hidden rows.
// Use Cases: Validate that a required code appears in the visible portion of a filtered report column. | Extract the first visible product identifier from a large worksheet where some rows are hidden. | Automate data quality checks that must ignore hidden rows when searching within a specific column range.
// AI Prompts: Show a compact Aspose.Cells C# snippet that searches K1:K500 for a string while automatically skipping hidden rows. | Explain how to modify the loop to collect all visible matches in K1:K500 instead of stopping at the first one. | Suggest an alternative using FindAll with a custom filter to exclude hidden rows in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFindIgnoreHidden
{
    // C# example that loads a workbook, sets a FindOptions range to K1:K500, and iteratively searches for a target string. Each match is checked with Cells.IsRowHidden so only visible rows are considered. The code reports the first visible occurrence or indicates that none was found.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create FindOptions and set the search range K1:K500
            FindOptions findOptions = new FindOptions();
            CellArea range = new CellArea
            {
                StartRow = 0,          // Row 1 (zero‑based)
                StartColumn = 10,      // Column K (zero‑based)
                EndRow = 499,          // Row 500
                EndColumn = 10         // Column K
            };
            findOptions.SetRange(range);
            findOptions.LookInType = LookInType.Values; // Search in cell values

            // The value to search for
            string searchValue = "TargetText";

            // Perform the search while skipping hidden rows
            Cell previousCell = null;
            Cell foundCell = null;

            while (true)
            {
                // Find the next occurrence using the previous cell as the start point
                foundCell = cells.Find(searchValue, previousCell, findOptions);
                if (foundCell == null)
                {
                    // No more matches
                    break;
                }

                // Check if the row containing the found cell is hidden
                if (!cells.IsRowHidden(foundCell.Row))
                {
                    // Found a visible row – exit the loop
                    break;
                }

                // Row is hidden – continue searching from this cell
                previousCell = foundCell;
            }

            if (foundCell != null && !cells.IsRowHidden(foundCell.Row))
            {
                Console.WriteLine($"Found '{searchValue}' in visible row at {foundCell.Name}");
            }
            else
            {
                Console.WriteLine($"Value '{searchValue}' not found in any visible rows within K1:K500.");
            }

            // Save the workbook if any modifications were made
            workbook.Save("output.xlsx");
        }
    }
}
