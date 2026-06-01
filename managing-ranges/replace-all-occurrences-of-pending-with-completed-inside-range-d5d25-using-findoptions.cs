using System;
using Aspose.Cells;

class ReplacePendingWithCompleted
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") if needed

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the search range D5:D25 (zero‑based indices)
        CellArea searchArea = new CellArea
        {
            StartRow = 4,   // Row 5
            StartColumn = 3, // Column D
            EndRow = 24,    // Row 25
            EndColumn = 3   // Column D
        };

        // Configure FindOptions to search only within the defined range
        FindOptions findOptions = new FindOptions
        {
            LookInType = LookInType.Values,          // search cell values
            LookAtType = LookAtType.EntireContent,   // match the whole cell content
        };
        findOptions.SetRange(searchArea);

        // Find the first occurrence of "Pending" within the range
        Cell foundCell = worksheet.Cells.Find("Pending", null, findOptions);

        // Replace all found occurrences with "Completed"
        while (foundCell != null)
        {
            foundCell.PutValue("Completed");

            // Continue searching from the cell just processed
            foundCell = worksheet.Cells.Find("Pending", foundCell, findOptions);
        }

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}