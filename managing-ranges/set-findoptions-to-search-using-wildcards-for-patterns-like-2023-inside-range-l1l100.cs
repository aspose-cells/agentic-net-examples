using System;
using Aspose.Cells;

class FindWithWildcards
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample data in column L (index 11) for demonstration
        cells["L1"].PutValue("Report_2022");
        cells["L2"].PutValue("Summary_2023_Q1");
        cells["L3"].PutValue("Data_2023_Final");
        cells["L4"].PutValue("Archive_2021");

        // Configure FindOptions to use Excel wildcards
        FindOptions findOptions = new FindOptions
        {
            // Search only in cell values (not formulas, comments, etc.)
            LookInType = LookInType.Values,
            // The whole cell content must match the wildcard pattern
            LookAtType = LookAtType.EntireContent,
            // Ensure the search key is treated as a wildcard pattern, not a regex
            RegexKey = false
        };

        // Define the search range L1:L100 (rows 0‑99, column 11)
        CellArea searchArea = new CellArea
        {
            StartRow = 0,          // Row 1
            StartColumn = 11,      // Column L (0‑based index)
            EndRow = 99,           // Row 100
            EndColumn = 11
        };
        findOptions.SetRange(searchArea);

        // Perform the search using the wildcard pattern "*2023*"
        Cell foundCell = cells.Find("*2023*", null, findOptions);

        // Output the result
        if (foundCell != null)
        {
            Console.WriteLine($"Found cell at {foundCell.Name} with value: {foundCell.StringValue}");
        }
        else
        {
            Console.WriteLine("No matching cell found in the specified range.");
        }

        // Save the workbook (optional)
        workbook.Save("FindWithWildcards.xlsx");
    }
}