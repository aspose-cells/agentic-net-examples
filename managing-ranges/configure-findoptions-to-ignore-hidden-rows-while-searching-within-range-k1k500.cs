using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the search range K1:K500 (column index 9, rows 0‑499)
        CellArea searchRange = new CellArea
        {
            StartRow = 0,
            StartColumn = 9,
            EndRow = 499,
            EndColumn = 9
        };

        // Configure FindOptions with the defined range
        FindOptions findOptions = new FindOptions();
        findOptions.SetRange(searchRange);
        findOptions.LookInType = LookInType.Values;          // search cell values
        findOptions.LookAtType = LookAtType.EntireContent;   // exact match (adjust as needed)

        // The value we want to find
        string whatToFind = "Target";

        // Perform the first search
        Cell foundCell = cells.Find(whatToFind, null, findOptions);

        // Skip hidden rows: keep searching until a visible row is found or no more matches
        while (foundCell != null && cells.IsRowHidden(foundCell.Row))
        {
            // Continue searching from the previously found cell
            foundCell = cells.Find(whatToFind, foundCell, findOptions);
        }

        if (foundCell != null)
        {
            Console.WriteLine($"Found at {foundCell.Name} (row {foundCell.Row + 1})");
        }
        else
        {
            Console.WriteLine("Value not found in visible rows within K1:K500.");
        }

        // Save the workbook if any modifications were made
        workbook.Save("output.xlsx");
    }
}