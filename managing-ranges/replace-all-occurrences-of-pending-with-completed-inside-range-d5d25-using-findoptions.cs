using System;
using Aspose.Cells;

namespace ReplacePendingWithCompleted
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") if needed
            Worksheet sheet = workbook.Worksheets[0];

            // Define the search range D5:D25 (zero‑based indexes)
            CellArea range = new CellArea
            {
                StartRow = 4,   // Row 5
                StartColumn = 3, // Column D
                EndRow = 24,    // Row 25
                EndColumn = 3   // Column D
            };

            // Configure FindOptions to limit the search to the defined range
            FindOptions findOptions = new FindOptions();
            findOptions.SetRange(range);
            findOptions.LookInType = LookInType.Values;          // search cell values
            findOptions.LookAtType = LookAtType.EntireContent;   // match whole cell content

            // Iterate through all cells that contain "Pending" within the range
            Cell previousCell = null;
            Cell foundCell = sheet.Cells.Find("Pending", previousCell, findOptions);
            while (foundCell != null)
            {
                // Replace the cell's value
                foundCell.PutValue("Completed");

                // Continue searching from the current cell
                previousCell = foundCell;
                foundCell = sheet.Cells.Find("Pending", previousCell, findOptions);
            }

            // Save the workbook
            workbook.Save("output.xlsx");
        }
    }
}