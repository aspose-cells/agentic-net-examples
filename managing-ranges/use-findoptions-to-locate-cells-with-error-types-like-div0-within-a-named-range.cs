// Title: Find #DIV/0! Errors in a Named Range using Aspose.Cells FindOptions (C#)
// Description: Demonstrates how to create a workbook, add formulas that generate #DIV/0! errors, define a named range, build a matching CellArea, configure FindOptions to search cell values, and iterate with Cells.Find to list every cell containing the division‑by‑zero error inside that range.
// Keywords: Aspose.Cells FindOptions error search | C# locate #DIV/0! cells | search Excel error values Aspose | named range cell lookup Aspose.Cells | find Excel error strings C#
// Common Searches: Aspose.Cells FindOptions find #DIV/0! in named range | C# search for Excel error values within a specific range | How to locate cells with #DIV/0! using Aspose.Cells | Find error cells in a named range Aspose.Cells .NET
// Developer Intent: Identify every cell that contains the #DIV/0! error inside a predefined named range.
// Use Cases: Generate a report of all division‑by‑zero errors for data‑quality audits. | Apply conditional formatting (e.g., red fill) to each error cell automatically. | Collect cell addresses into a list for downstream processing or logging.
// AI Prompts: Create code that highlights each #DIV/0! cell with a red background instead of printing its address. | Show how to return a List<string> of the error cell references from a helper method. | Explain how to modify FindOptions to detect other Excel errors such as #N/A, #VALUE!, or #REF!.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a workbook, add formulas that generate #DIV/0! errors, define a named range, build a matching CellArea, configure FindOptions to search cell values, and iterate with Cells.Find to list every cell containing the division‑by‑zero error inside that range.
class FindErrorCellsInNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add formulas that will generate #DIV/0! errors
            worksheet.Cells["A1"].Formula = "=1/0";          // #DIV/0! error
            worksheet.Cells["A2"].Formula = "=B2";          // No error (B2 is empty)
            worksheet.Cells["B1"].Formula = "=SUM(1,2)";    // Normal value
            worksheet.Cells["B2"].Formula = "=A1";          // Propagates #DIV/0! error

            // Calculate formulas so that error values are materialized
            workbook.CalculateFormula();

            // Define a named range that covers the area we want to search
            worksheet.Cells.CreateRange("A1", "B2").Name = "ErrorRange";

            // Retrieve the named range as a Range object (use alias to avoid conflict with System.Range)
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("ErrorRange");

            // Build a CellArea that represents the same area as the named range
            CellArea searchArea = new CellArea
            {
                StartRow = namedRange.FirstRow,
                StartColumn = namedRange.FirstColumn,
                EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
            };

            // Configure FindOptions to search within the defined range and look at cell values
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,          // Search in cell values (including error strings)
                LookAtType = LookAtType.EntireContent    // Exact match
            };
            findOptions.SetRange(searchArea);            // Apply the search range

            // The error string that represents a division-by-zero error in Excel
            const string errorString = "#DIV/0!";

            // Iterate through all cells that contain the error string within the named range
            Cell? previousCell = null;
            while (true)
            {
                Cell foundCell = worksheet.Cells.Find(errorString, previousCell, findOptions);
                if (foundCell == null)
                    break;

                Console.WriteLine($"Error cell found at {foundCell.Name}");
                previousCell = foundCell; // Continue searching after the current cell
            }

            // Save the workbook (optional, just to demonstrate lifecycle compliance)
            workbook.Save("FindErrorCellsInNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
