using System;
using Aspose.Cells;

namespace AsposeCellsFindIgnoreHiddenRows
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data in column K (index 10)
            for (int i = 0; i < 500; i++)
            {
                cells[i, 10].PutValue(i % 50 == 0 ? "Target" : $"Value {i}");
            }

            // Hide a few rows that also contain the target value
            // Row indices are zero‑based; hide rows 0, 100 and 200
            cells.HideRows(0, 1);
            cells.HideRows(100, 1);
            cells.HideRows(200, 1);

            // Configure FindOptions to search only within K1:K500
            FindOptions findOptions = new FindOptions();
            CellArea searchArea = new CellArea
            {
                StartRow = 0,          // K1 -> row 0
                StartColumn = 10,      // Column K -> index 10
                EndRow = 499,          // K500 -> row 499
                EndColumn = 10
            };
            findOptions.SetRange(searchArea);

            // Optional: set other search preferences
            findOptions.LookInType = LookInType.Values;
            findOptions.LookAtType = LookAtType.EntireContent;
            findOptions.SearchOrderByRows = true;

            // Perform the search, skipping hidden rows
            Cell previousCell = null;
            Cell foundCell = null;

            while (true)
            {
                // Find the next occurrence of "Target"
                foundCell = cells.Find("Target", previousCell, findOptions);

                // No more matches
                if (foundCell == null)
                    break;

                // If the row is not hidden, we have our result
                if (!cells.IsRowHidden(foundCell.Row))
                    break;

                // Otherwise continue searching after this hidden cell
                previousCell = foundCell;
            }

            if (foundCell != null && !cells.IsRowHidden(foundCell.Row))
                Console.WriteLine($"Found visible 'Target' at {foundCell.Name}");
            else
                Console.WriteLine("Visible 'Target' not found in the specified range.");

            // Save the workbook (demonstrates usage of the required save rule)
            workbook.Save("FindIgnoreHiddenRowsResult.xlsx");
        }
    }
}