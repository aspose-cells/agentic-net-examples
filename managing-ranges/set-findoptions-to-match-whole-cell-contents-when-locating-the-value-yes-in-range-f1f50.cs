// Title: C# Aspose.Cells FindOptions – Find whole‑cell value "Yes" in range F1:F50
// Description: Demonstrates how to configure FindOptions with LookInType.Values and LookAtType.EntireContent, limit the search to the cell area F1:F50, and use Worksheet.Cells.Find to retrieve the first cell whose complete content equals "Yes". The example prints the cell address and optionally saves the workbook.
// Keywords: Aspose.Cells FindOptions C# | search whole cell value | LookAtType.EntireContent | Find exact text in range | Worksheet.Cells.Find example | limit search to F1:F50 | C# spreadsheet exact match
// Common Searches: Aspose.Cells find exact cell value C# | How to use FindOptions for whole‑cell match | Search for "Yes" in column F with Aspose.Cells | Limit Aspose.Cells Find to a specific range | LookAtType.EntireContent usage example
// Developer Intent: Retrieve the first cell whose entire content is exactly "Yes" within the range F1:F50 using Aspose.Cells in C#.
// Use Cases: Validate that a status column contains only the exact word "Yes" before processing rows. | Trigger business logic only when a flag column matches the whole‑cell value "Yes". | Generate a filtered report by locating rows with an exact keyword in a specific column.
// AI Prompts: Show how to return all cells with the exact value "Yes" in F1:F50 using FindAll. | Provide a sample that searches formulas for an exact match with LookInType.Formulas. | Explain how to combine multiple FindOptions searches to locate several exact strings in the same column.

using System;
using Aspose.Cells;

// Demonstrates how to configure FindOptions with LookInType.Values and LookAtType.EntireContent, limit the search to the cell area F1:F50, and use Worksheet.Cells.Find to retrieve the first cell whose complete content equals "Yes". The example prints the cell address and optionally saves the workbook.
class FindWholeCellContentExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Example data (optional, can be omitted if workbook already has data)
        worksheet.Cells["F1"].PutValue("Yes");
        worksheet.Cells["F2"].PutValue("Yes, indeed");
        worksheet.Cells["F3"].PutValue("No");

        // Configure FindOptions to match the entire cell content
        FindOptions findOptions = new FindOptions
        {
            LookInType = LookInType.Values,          // Search in cell values
            LookAtType = LookAtType.EntireContent    // Match whole cell contents
        };

        // Define the search range F1:F50 (zero‑based indices)
        CellArea searchArea = new CellArea
        {
            StartRow = 0,      // Row 1
            StartColumn = 5,   // Column F
            EndRow = 49,       // Row 50
            EndColumn = 5      // Column F
        };
        findOptions.SetRange(searchArea);

        // Perform the search for the exact string "Yes"
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

        // Save the workbook (optional)
        workbook.Save("FindResult.xlsx");
    }
}
