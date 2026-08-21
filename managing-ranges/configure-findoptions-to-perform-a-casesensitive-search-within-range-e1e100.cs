// Title: Case‑Sensitive Find in Column E (E1:E100) using Aspose.Cells for .NET
// Description: Creates a workbook, populates cells E1‑E5, sets FindOptions.CaseSensitive = true, defines a CellArea covering rows 1‑100 of column E, and uses Worksheet.Cells.Find to locate an exact‑case match for "Apple". The result is printed and the workbook saved.
// Keywords: Aspose.Cells | .NET | C# | FindOptions | case sensitive search | Excel column E | E1:E100 range | CellArea | Worksheet.Cells.Find | programmatic Excel lookup
// Common Searches: Aspose.Cells case sensitive FindOptions example | search specific column range C# Aspose.Cells | find text in Excel column E with case sensitivity | how to set search range for Find in Aspose.Cells | C# case‑sensitive lookup in Excel using Aspose
// Developer Intent: Find the first cell that exactly matches a given string in column E while preserving case.
// Use Cases: Validate that product codes entered with exact casing exist in a catalog column. | Extract rows where a case‑sensitive identifier appears in a specific worksheet column. | Enforce case‑sensitive data quality rules during automated Excel processing.
// AI Prompts: Show how to retrieve all case‑sensitive matches in E1:E100 instead of only the first one. | Demonstrate using FindOptions with a regular expression for case‑sensitive searches. | Adapt the code to search for multiple case‑sensitive strings within the same column range.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates cells E1‑E5, sets FindOptions.CaseSensitive = true, defines a CellArea covering rows 1‑100 of column E, and uses Worksheet.Cells.Find to locate an exact‑case match for "Apple". The result is printed and the workbook saved.
    class FindCaseSensitiveInRange
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data in column E (E1:E5 for demonstration)
            worksheet.Cells["E1"].PutValue("Apple");
            worksheet.Cells["E2"].PutValue("apple");
            worksheet.Cells["E3"].PutValue("Banana");
            worksheet.Cells["E4"].PutValue("APPLE");
            worksheet.Cells["E5"].PutValue("Orange");

            // Configure FindOptions for case‑sensitive search
            FindOptions findOptions = new FindOptions
            {
                CaseSensitive = true // enable case sensitivity
            };

            // Define the search range E1:E100
            CellArea searchRange = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 4,   // Column E (zero‑based)
                EndRow = 99,       // Row 100
                EndColumn = 4      // Column E
            };
            findOptions.SetRange(searchRange);

            // Perform the search for the string "Apple"
            Cell foundCell = worksheet.Cells.Find("Apple", null, findOptions);

            // Output the result
            Console.WriteLine(foundCell == null
                ? "Not found (case‑sensitive)"
                : $"Found at: {foundCell.Name}");

            // Save the workbook (optional)
            string outputPath = "FindCaseSensitiveInRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
