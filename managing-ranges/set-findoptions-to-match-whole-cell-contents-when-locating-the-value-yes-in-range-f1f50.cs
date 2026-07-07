using System;
using Aspose.Cells;

namespace AsposeCellsFindExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Example data: populate some cells in column F (index 5)
            worksheet.Cells["F1"].PutValue("Yes");
            worksheet.Cells["F2"].PutValue("No");
            worksheet.Cells["F3"].PutValue("Yes ");
            worksheet.Cells["F4"].PutValue("Maybe");
            // ... populate as needed ...

            // Define the search range F1:F50
            CellArea searchArea = new CellArea
            {
                StartRow = 0,      // Row 1 (0‑based)
                StartColumn = 5,   // Column F (0‑based)
                EndRow = 49,       // Row 50
                EndColumn = 5      // Column F
            };

            // Configure FindOptions to match the entire cell contents
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,          // Search in cell values
                LookAtType = LookAtType.EntireContent    // Exact whole‑cell match
            };
            findOptions.SetRange(searchArea); // Apply the range to the options

            // Perform the search for the exact value "Yes"
            Cell foundCell = worksheet.Cells.Find("Yes", null, findOptions);

            // Output the result
            if (foundCell != null)
            {
                Console.WriteLine($"Found \"Yes\" at cell {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("The value \"Yes\" was not found in the specified range.");
            }

            // Optionally save the workbook (if you need to inspect it)
            // workbook.Save("FindResult.xlsx");
        }
    }
}