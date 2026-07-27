using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate column I (index 8) with sample data, inserting "Error" at some rows
        for (int i = 0; i < 200; i++)
        {
            if (i == 50 || i == 150)
                worksheet.Cells[i, 8].PutValue("Error");
            else
                worksheet.Cells[i, 8].PutValue($"Value_{i}");
        }

        // Define the search range I1:I200
        CellArea searchRange = new CellArea
        {
            StartRow = 0,
            StartColumn = 8,
            EndRow = 199,
            EndColumn = 8
        };

        // Configure FindOptions for a backward search
        FindOptions options = new FindOptions
        {
            LookInType = LookInType.Values,          // Search in cell values
            LookAtType = LookAtType.EntireContent,   // Exact match
            SearchBackward = true                    // Search from bottom to top
        };
        options.SetRange(searchRange);

        // Perform the backward search for the string "Error"
        Cell foundCell = worksheet.Cells.Find("Error", null, options);

        if (foundCell != null)
            Console.WriteLine($"Found 'Error' at {foundCell.Name}");
        else
            Console.WriteLine("No 'Error' found in the specified range.");

        // Save the workbook (optional)
        workbook.Save("BackwardSearchResult.xlsx");
    }
}