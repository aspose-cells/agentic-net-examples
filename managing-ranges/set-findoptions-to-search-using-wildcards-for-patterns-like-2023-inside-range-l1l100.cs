using System;
using Aspose.Cells;

namespace AsposeCellsFindWildcardExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column L (index 11) for demonstration
            for (int i = 0; i < 100; i++)
            {
                // Example values, some contain the year 2023
                string value = (i % 10 == 0) ? $"Report_2023_Q{i / 10}" : $"Data_{i}";
                cells[i, 11].PutValue(value); // Row i, Column L
            }

            // Create FindOptions instance
            FindOptions findOptions = new FindOptions();

            // Define the search range L1:L100
            CellArea searchArea = new CellArea
            {
                StartRow = 0,          // Row 1 (0‑based)
                StartColumn = 11,      // Column L (0‑based)
                EndRow = 99,           // Row 100
                EndColumn = 11         // Column L
            };
            findOptions.SetRange(searchArea);

            // Configure options to use wildcards
            findOptions.LookInType = LookInType.Values;   // Search in cell values
            findOptions.LookAtType = LookAtType.Contains; // Allows wildcard handling
            findOptions.RegexKey = false;                 // Use Excel‑style wildcards, not regex

            // Search for the pattern "*2023*"
            Cell foundCell = cells.Find("*2023*", null, findOptions);

            if (foundCell != null)
            {
                Console.WriteLine($"Found cell at {foundCell.Name} with value: {foundCell.StringValue}");
            }
            else
            {
                Console.WriteLine("No cell matching the pattern '*2023*' was found in the specified range.");
            }

            // Optionally save the workbook (demonstration purpose)
            workbook.Save("FindWildcardResult.xlsx");
        }
    }
}