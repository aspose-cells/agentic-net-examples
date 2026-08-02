// Title: Copy a cell range to CSV with delimiters and text qualifiers using Aspose.Cells for .NET
// Description: Creates a workbook, fills cells A1:B3 (including a comma‑containing value), defines a CellArea for that range, and saves it as a CSV file using TxtSaveOptions with a comma separator and always‑quoted fields.
// Keywords: Aspose.Cells CSV export | C# range to CSV | TxtSaveOptions delimiter | always quote CSV fields | export selected cells Aspose | .NET workbook to CSV
// Common Searches: Aspose.Cells export selected range to CSV | C# save worksheet area as CSV with quotes | how to preserve commas in CSV using Aspose | TxtSaveOptions QuoteType Always example | export part of Excel sheet to CSV .NET
// Developer Intent: Save only the A1:B3 range as a CSV file while ensuring commas inside values are retained and every field is enclosed in quotes.
// Use Cases: Generating a CSV report that includes only specific rows and columns. | Creating data files for systems that require all fields to be quoted, even when delimiters appear in the content. | Extracting a subset of a large workbook for lightweight data exchange.
// AI Prompts: Write C# code with Aspose.Cells to export a defined cell area to CSV using a custom delimiter and always‑quote option. | Explain how TxtSaveOptions Separator, QuoteType, and ExportArea work together to preserve delimiters and text qualifiers in a CSV export.

using System;
using Aspose.Cells;

namespace AsposeCellsRangeToCsv
{
    // Creates a workbook, fills cells A1:B3 (including a comma‑containing value), defines a CellArea for that range, and saves it as a CSV file using TxtSaveOptions with a comma separator and always‑quoted fields.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including a value that contains the delimiter)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John, Doe"); // comma inside the text
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);

            // Define the range to be exported (A1:B3)
            CellArea exportArea = new CellArea
            {
                StartRow = 0,   // Row index is zero‑based
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 1
            };

            // Configure text save options for CSV export
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ',',                     // Use comma as delimiter
                QuoteType = TxtValueQuoteType.Always, // Always quote fields (preserves text qualifiers)
                ExportArea = exportArea               // Export only the defined range
            };

            // Save the selected range to a CSV file
            workbook.Save("ExportedRange.csv", saveOptions);
        }
    }
}
