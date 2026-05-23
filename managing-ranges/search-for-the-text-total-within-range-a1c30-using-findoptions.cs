using Aspose.Cells;
using System;

class FindTotalInRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional, can be omitted if workbook is loaded from a file)
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Total");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(20);
        sheet.Cells["A3"].PutValue("Total"); // another occurrence

        // Configure FindOptions to search within A1:C30
        FindOptions options = new FindOptions();
        CellArea searchArea = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            StartColumn = 0,   // Column A (zero‑based)
            EndRow = 29,       // Row 30 (zero‑based)
            EndColumn = 2      // Column C (zero‑based)
        };
        options.SetRange(searchArea);
        options.LookInType = LookInType.Values;      // Search cell values
        options.LookAtType = LookAtType.Contains;    // Partial match (default)

        // Perform the search for the text "Total"
        Cell foundCell = sheet.Cells.Find("Total", null, options);

        if (foundCell != null)
        {
            Console.WriteLine($"Found 'Total' at {foundCell.Name} (Row {foundCell.Row}, Column {foundCell.Column})");
        }
        else
        {
            Console.WriteLine("Text 'Total' not found in the specified range.");
        }

        // Save the workbook (optional)
        workbook.Save("FindTotalResult.xlsx");
    }
}