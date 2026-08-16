// Title: Aspose.Cells C# – Case‑Insensitive FindOptions Search for "FY" in a Defined Cell Range
// Description: This example creates a workbook, fills cells A1:C5 with sample text, defines a CellArea covering that range, configures FindOptions to ignore case and match cells that contain the key, and uses Worksheet.Cells.Find to locate the first occurrence of the abbreviation "FY" regardless of case. The result prints the cell address and value, then saves the workbook.
// Keywords: Aspose.Cells FindOptions | case insensitive search C# | search abbreviation FY | CellArea range search | LookAtType.Contains | Aspose.Cells .NET example | worksheet cell find | Excel automation
// Common Searches: Aspose.Cells case insensitive FindOptions example | How to search for "FY" in a specific range with Aspose.Cells | Find cells containing a substring in Aspose.Cells .NET | Use CellArea with Find method in Aspose.Cells | Lookup abbreviation in Excel using C# Aspose.Cells
// Developer Intent: Locate the first cell that contains the abbreviation "FY" (any case) within the A1:C5 range of a worksheet.
// Use Cases: Detect fiscal‑year abbreviations before generating financial reports. | Validate required abbreviations in imported data sheets to avoid processing errors. | Highlight cells with "FY" for conditional formatting or downstream transformations.
// AI Prompts: Show how to retrieve all cells that contain "FY" instead of just the first match. | Explain how to change FindOptions for an exact whole‑cell match rather than a contains search. | Provide a sample that searches formulas (LookInType.Formulas) for the abbreviation "FY".

using System;
using Aspose.Cells;

namespace FindAbbreviationExample
{
    // This example creates a workbook, fills cells A1:C5 with sample text, defines a CellArea covering that range, configures FindOptions to ignore case and match cells that contain the key, and uses Worksheet.Cells.Find to locate the first occurrence of the abbreviation "FY" regardless of case. The result prints the cell address and value, then saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data within the range A1:C5
            worksheet.Cells["A1"].PutValue("FY");
            worksheet.Cells["A2"].PutValue("fy");
            worksheet.Cells["B3"].PutValue("Fiscal Year");
            worksheet.Cells["C4"].PutValue("fy2021");
            worksheet.Cells["A5"].PutValue("Other");

            // Define the search range (A1:C5)
            CellArea searchArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 4,        // Row 5
                EndColumn = 2      // Column C
            };

            // Configure FindOptions for a case‑insensitive search
            FindOptions options = new FindOptions
            {
                CaseSensitive = false,                 // Ignore case
                LookInType = LookInType.Values,        // Search cell values
                LookAtType = LookAtType.Contains        // Match if the cell contains the key
            };
            options.SetRange(searchArea);               // Apply the defined range

            // Perform the search for the abbreviation "FY"
            Cell foundCell = worksheet.Cells.Find("FY", null, options);

            // Output the result
            if (foundCell != null)
            {
                Console.WriteLine($"Found \"FY\" at cell {foundCell.Name} with value \"{foundCell.StringValue}\".");
            }
            else
            {
                Console.WriteLine("The abbreviation \"FY\" was not found in the specified range.");
            }

            // Save the workbook (optional)
            workbook.Save("FindAbbreviationResult.xlsx");
        }
    }
}
