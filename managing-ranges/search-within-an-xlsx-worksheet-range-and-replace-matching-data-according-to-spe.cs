using System;
using Aspose.Cells;

namespace AsposeCellsSearchAndReplace
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the search range: rows 1-10 (0‑based index 0‑9), columns A‑E (0‑4)
            CellArea searchRange = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 4
            };

            // Configure find options to search within the defined range,
            // look at cell values, case‑insensitive, and match if the value contains the key
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                CaseSensitive = false
            };
            findOptions.SetRange(searchRange);

            // The text to find and its replacement
            string placeholder = "OldText";
            string newValue = "NewText";

            // Prepare replace options (non‑case‑sensitive, partial match allowed)
            ReplaceOptions replaceOptions = new ReplaceOptions
            {
                CaseSensitive = false,
                MatchEntireCellContents = false
            };

            // Start searching from the beginning of the range
            Cell previousCell = null;
            Cell foundCell = cells.Find(placeholder, previousCell, findOptions);

            // Loop through all occurrences within the range
            while (foundCell != null)
            {
                // Replace the placeholder in the found cell using the specified options
                foundCell.Replace(placeholder, newValue, replaceOptions);

                // Continue searching after the current cell
                previousCell = foundCell;
                foundCell = cells.Find(placeholder, previousCell, findOptions);
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}