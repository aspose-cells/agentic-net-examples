// Title: Find cells containing "Total" in range A1:C30 using Aspose.Cells FindOptions (C#)
// Description: Creates a workbook, optionally fills sample data, defines a CellArea for A1:C30, configures FindOptions to search cell values with a Contains match, and uses Worksheet.Cells.Find to locate the first occurrence of "Total" within the specified range, then outputs the address and saves the file.
// Keywords: Aspose.Cells | FindOptions | C# | search range | find text | Excel worksheet | CellArea | Find method | contains match | Total label
// Common Searches: Aspose.Cells find text in specific range C# | How to use FindOptions to locate "Total" in Excel | Search cells A1 to C30 for a substring with Aspose.Cells | C# code to find first occurrence of a word in a worksheet area
// Developer Intent: Locate the first cell that contains the word "Total" inside the A1:C30 range of a worksheet using Aspose.Cells for .NET.
// Use Cases: Validate that a generated report includes a "Total" label within the expected rows before export. | Retrieve the row and column of the "Total" cell to drive subsequent calculations or data aggregation. | Apply conditional formatting or data validation to any cell that contains the word "Total" after it is found.
// AI Prompts: Generate C# code with Aspose.Cells that returns all cell addresses containing "Total" in the range A1:C30. | Explain how to modify FindOptions for a case‑insensitive search of "total" across an entire worksheet. | Show how to iterate through multiple matches of "Total" using FindNext with the same FindOptions configuration.

using System;
using Aspose.Cells;

namespace AsposeCellsFindTotalDemo
{
    // Creates a workbook, optionally fills sample data, defines a CellArea for A1:C30, configures FindOptions to search cell values with a Contains match, and uses Worksheet.Cells.Find to locate the first occurrence of "Total" within the specified range, then outputs the address and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Populate some data for demonstration
            worksheet.Cells["A1"].PutValue("Item");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["C1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Total");
            worksheet.Cells["B2"].PutValue(150);
            worksheet.Cells["C2"].PutValue(12.5);
            worksheet.Cells["A10"].PutValue("Grand Total");
            worksheet.Cells["B10"].PutValue(300);
            worksheet.Cells["C10"].PutValue(25.0);

            // Configure find options
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,      // Search in cell values
                LookAtType = LookAtType.Contains,    // Match if the cell contains the search text
                SearchBackward = false,              // Search forward
                SearchOrderByRows = true             // Row‑wise search
            };

            // Define the search range A1:C30 (zero‑based indices)
            CellArea searchArea = new CellArea
            {
                StartRow = 0,    // Row 1
                StartColumn = 0, // Column A
                EndRow = 29,     // Row 30
                EndColumn = 2    // Column C
            };
            findOptions.SetRange(searchArea);

            // Perform the search for the text "Total"
            Cell foundCell = worksheet.Cells.Find("Total", null, findOptions);

            // Output the result
            if (foundCell != null)
            {
                Console.WriteLine($"Found \"Total\" at cell {foundCell.Name} (Row {foundCell.Row + 1}, Column {foundCell.Column + 1})");
            }
            else
            {
                Console.WriteLine("The text \"Total\" was not found within the specified range.");
            }

            // Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save("FindTotalDemo.xlsx");
        }
    }
}
