using Aspose.Cells;
using System;

class FindTotalInRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional, can be omitted if workbook already has data)
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Total");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(20);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(5);
        sheet.Cells["C3"].PutValue(15);

        // Define the search range A1:C30
        CellArea searchRange = new CellArea
        {
            StartRow = 0,      // Row index for A1
            StartColumn = 0,   // Column index for A
            EndRow = 29,       // Row index for row 30 (0‑based)
            EndColumn = 2      // Column index for C
        };

        // Configure FindOptions
        FindOptions options = new FindOptions();
        options.SetRange(searchRange);
        options.LookInType = LookInType.Values;      // Search in cell values
        options.LookAtType = LookAtType.Contains;    // Look for cells containing the text

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