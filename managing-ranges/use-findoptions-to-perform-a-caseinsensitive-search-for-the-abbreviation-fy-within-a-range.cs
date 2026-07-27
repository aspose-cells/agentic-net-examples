using System;
using Aspose.Cells;

class FindFYExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with different case variations of "FY"
        sheet.Cells["A1"].PutValue("FY");
        sheet.Cells["A2"].PutValue("fy");
        sheet.Cells["A3"].PutValue("FYear");
        sheet.Cells["B1"].PutValue("Other");
        sheet.Cells["C5"].PutValue("fy");

        // Define the search range (A1:C5)
        CellArea range = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            StartColumn = 0,   // Column A
            EndRow = 4,        // Row 5
            EndColumn = 2      // Column C
        };

        // Configure FindOptions for a case‑insensitive search
        FindOptions options = new FindOptions
        {
            CaseSensitive = false,               // ignore case
            LookInType = LookInType.Values,      // search cell values
            LookAtType = LookAtType.EntireContent // exact match of the whole cell content
        };
        options.SetRange(range); // limit the search to the defined range

        // Perform the search for the abbreviation "FY"
        Cell foundCell = sheet.Cells.Find("FY", null, options);

        // Output the result
        if (foundCell != null)
        {
            Console.WriteLine($"Found at {foundCell.Name} with value '{foundCell.StringValue}'");
        }
        else
        {
            Console.WriteLine("Not found");
        }

        // Save the workbook to verify the data (optional)
        workbook.Save("FindFYResult.xlsx");
    }
}