// Title: C# – Limit Aspose.Cells Find to G1:G10 and H1:H10 with FindOptions.SetRange
// Description: Creates a workbook, fills columns G and H (rows 1‑10) with sample data, places the string "Target" in G5 and H8, defines a CellArea covering G1:H10, applies it to FindOptions via SetRange, and searches for "Target" only within that area, returning the cell address if found.
// Keywords: Aspose.Cells | FindOptions | SetRange | CellArea | search specific range | C# example | non‑contiguous columns | G1:G10 | H1:H10
// Common Searches: Aspose.Cells limit Find to certain columns | SetRange for multiple columns .NET | Search only G1:G10 and H1:H10 with Aspose.Cells | FindOptions CellArea example C# | How to restrict Aspose.Cells Find to a custom range
// Developer Intent: Search for a value exclusively inside the G1:G10 and H1:H10 ranges using Aspose.Cells.
// Use Cases: Locate a keyword in two adjacent columns without scanning the entire sheet. | Validate that user input appears only in designated columns before processing. | Generate a report of matches confined to specific columns.
// AI Prompts: Write C# code that searches for a string in G1:G10 and H1:H10 using FindOptions.SetRange. | Explain how to define a CellArea that covers multiple columns for an Aspose.Cells search. | Show how to handle a missing target value when using Find with a limited range.

using System;
using Aspose.Cells;

namespace AsposeCellsSearchNonContiguous
{
    // Creates a workbook, fills columns G and H (rows 1‑10) with sample data, places the string "Target" in G5 and H8, defines a CellArea covering G1:H10, applies it to FindOptions via SetRange, and searches for "Target" only within that area, returning the cell address if found.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate G1:G10 and H1:H10 with sample data
            for (int i = 0; i < 10; i++)
            {
                // Column G (index 6)
                cells[i, 6].PutValue($"G{i + 1}");
                // Column H (index 7)
                cells[i, 7].PutValue($"H{i + 1}");
            }

            // Put a target value that we will search for
            cells[4, 6].PutValue("Target"); // G5
            cells[7, 7].PutValue("Target"); // H8

            // Configure FindOptions to limit the search to G1:G10 and H1:H10.
            // Since these two columns form a single rectangular area (G1:H10),
            // we can set the range using a CellArea that covers both columns.
            FindOptions findOptions = new FindOptions();

            CellArea searchArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 6,   // Column G (zero‑based)
                EndRow = 9,        // Row 10
                EndColumn = 7      // Column H
            };

            findOptions.SetRange(searchArea);

            // Perform the search for the string "Target"
            Cell foundCell = cells.Find("Target", null, findOptions);

            if (foundCell != null)
            {
                Console.WriteLine($"Found 'Target' at {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("Value not found within the specified ranges.");
            }

            // Save the workbook (optional, just to visualize the data)
            workbook.Save("SearchNonContiguousDemo.xlsx");
        }
    }
}
