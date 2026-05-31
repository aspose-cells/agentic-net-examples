using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data in column E (E1:E5)
        worksheet.Cells["E1"].PutValue("Apple");
        worksheet.Cells["E2"].PutValue("apple");
        worksheet.Cells["E3"].PutValue("Banana");
        worksheet.Cells["E4"].PutValue("APPLE");
        worksheet.Cells["E5"].PutValue("Orange");

        // Configure FindOptions for a case‑sensitive search
        FindOptions findOptions = new FindOptions();
        findOptions.CaseSensitive = true; // enable case sensitivity

        // Define the search range E1:E100
        CellArea searchRange = new CellArea();
        searchRange.StartRow = 0;      // Row 1 (zero‑based index)
        searchRange.StartColumn = 4;   // Column E (zero‑based index)
        searchRange.EndRow = 99;       // Row 100
        searchRange.EndColumn = 4;     // Column E
        findOptions.SetRange(searchRange);

        // Perform the search for the string "Apple"
        Cell foundCell = worksheet.Cells.Find("Apple", null, findOptions);

        // Output the result
        Console.WriteLine(foundCell != null
            ? $"Found at {foundCell.Name}"
            : "Not found (case‑sensitive)");

        // Save the workbook (optional)
        workbook.Save("FindCaseSensitiveDemo.xlsx");
    }
}