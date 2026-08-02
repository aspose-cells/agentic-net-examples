using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data in the target range (optional)
        worksheet.Cells["E1"].PutValue("Apple");
        worksheet.Cells["E2"].PutValue("apple");
        worksheet.Cells["E3"].PutValue("APPLE");

        // Configure FindOptions for a case‑sensitive search
        FindOptions options = new FindOptions();
        options.CaseSensitive = true; // enable case sensitivity

        // Define the search range E1:E100
        CellArea range = new CellArea();
        range.StartRow = 0;      // Row 1 (zero‑based index)
        range.StartColumn = 4;   // Column E (zero‑based index)
        range.EndRow = 99;       // Row 100 (zero‑based index)
        range.EndColumn = 4;     // Column E
        options.SetRange(range);

        // Perform the search (example: looking for "Apple")
        Cell foundCell = worksheet.Cells.Find("Apple", null, options);

        // Output the result
        Console.WriteLine(foundCell != null ? $"Found at {foundCell.Name}" : "Not found");
    }
}