// Title: C# – Replace “Pending” with “Completed” in D5:D25 using Aspose.Cells FindOptions
// Description: Loads an Excel workbook, defines a CellArea for D5:D25, uses FindOptions to locate cells whose value equals “Pending”, replaces each with “Completed”, and saves the updated file.
// Keywords: Aspose.Cells FindOptions | C# find and replace Excel | replace text in specific range | cell area search Aspose | update status column Excel C# | Aspose.Cells replace values
// Common Searches: Aspose.Cells replace text in a range C# | How to limit FindOptions search to D5:D25 | C# find and replace specific cell values with Aspose.Cells | Replace all occurrences of a word in Excel using Aspose.Cells
// Developer Intent: Replace every occurrence of the word "Pending" with "Completed" within the D5:D25 range of an Excel worksheet.
// Use Cases: Change a status column from "Pending" to "Completed" after a batch job finishes. | Prepare data for reporting by converting placeholder values to final results in a defined range. | Cleanse imported spreadsheets by updating specific text within a limited cell area.
// AI Prompts: Show how to modify the script so it replaces "Pending" with "In Progress" only when column C contains a numeric ID. | Generate code that logs the address of each cell changed from "Pending" to "Completed" into a text file. | Provide an alternative implementation that uses Cells.Replace with a CellArea instead of FindOptions.

using System;
using Aspose.Cells;

namespace AsposeCellsReplacePending
{
    // Loads an Excel workbook, defines a CellArea for D5:D25, uses FindOptions to locate cells whose value equals “Pending”, replaces each with “Completed”, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure FindOptions to limit the search to the range D5:D25
            FindOptions findOptions = new FindOptions();
            CellArea range = new CellArea
            {
                StartRow = 4,    // Row index is zero‑based (D5 -> row 4)
                StartColumn = 3, // Column D -> index 3
                EndRow = 24,     // D25 -> row 24
                EndColumn = 3    // Same column D
            };
            findOptions.SetRange(range);
            findOptions.LookInType = LookInType.Values;   // Search cell values
            findOptions.LookAtType = LookAtType.EntireContent; // Exact match

            // Repeatedly find cells containing "Pending" within the defined range and replace them
            Cell foundCell = cells.Find("Pending", null, findOptions);
            while (foundCell != null)
            {
                // Replace the cell's value
                foundCell.PutValue("Completed");

                // Continue searching from the next cell
                foundCell = cells.Find("Pending", foundCell, findOptions);
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
