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

            // Populate sample data (including different cases of "FY")
            worksheet.Cells["A1"].PutValue("FY2021");
            worksheet.Cells["A2"].PutValue("fy2022");
            worksheet.Cells["A3"].PutValue("Fiscal Year");
            worksheet.Cells["B1"].PutValue("Q1");
            worksheet.Cells["B2"].PutValue("fy");
            worksheet.Cells["B3"].PutValue("FY");

            // Define the search range (A1:B3)
            CellArea searchArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 2,        // Row 3
                EndColumn = 1      // Column B
            };

            // Configure FindOptions for a case‑insensitive search
            FindOptions options = new FindOptions
            {
                CaseSensitive = false,          // Ignore case
                LookInType = LookInType.Values, // Search cell values
                LookAtType = LookAtType.Contains // Match if the cell contains the key
            };
            options.SetRange(searchArea);        // Apply the defined range

            // Perform the search for the abbreviation "FY"
            Cell foundCell = worksheet.Cells.Find("FY", null, options);

            // Output the result
            if (foundCell != null)
            {
                Console.WriteLine($"Found \"FY\" in cell {foundCell.Name} with value \"{foundCell.StringValue}\"");
            }
            else
            {
                Console.WriteLine("The abbreviation \"FY\" was not found in the specified range.");
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("FindFYResult.xlsx");
        }
    }
}