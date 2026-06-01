using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data in column F (optional, for demonstration)
        worksheet.Cells["F1"].PutValue("Yes");
        worksheet.Cells["F2"].PutValue("Yes ");
        worksheet.Cells["F3"].PutValue("No");
        worksheet.Cells["F4"].PutValue("Maybe");
        worksheet.Cells["F5"].PutValue("Yes");

        // Configure FindOptions to match the entire cell contents
        FindOptions findOptions = new FindOptions
        {
            LookInType = LookInType.Values,
            LookAtType = LookAtType.EntireContent // exact match
        };

        // Define the search range F1:F50
        CellArea searchArea = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based index)
            StartColumn = 5,   // Column F (zero‑based index)
            EndRow = 49,       // Row 50
            EndColumn = 5      // Column F
        };
        findOptions.SetRange(searchArea);

        // Perform the find operation for the value "Yes"
        Cell foundCell = worksheet.Cells.Find("Yes", null, findOptions);

        // Output the result
        if (foundCell != null)
        {
            Console.WriteLine($"Found 'Yes' at {foundCell.Name}");
        }
        else
        {
            Console.WriteLine("Value 'Yes' not found with whole cell match.");
        }

        // Save the workbook (optional)
        workbook.Save("FindWholeCellContent.xlsx");
    }
}