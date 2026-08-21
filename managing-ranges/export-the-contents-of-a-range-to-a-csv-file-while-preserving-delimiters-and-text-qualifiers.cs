// Title: Export a Selected Cell Range to CSV with Delimiters and Text Qualifiers – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to export a defined CellArea (A1:B3) to a CSV file using Aspose.Cells. The example sets TxtSaveOptions with a comma separator, Normal quote type, UTF‑8 encoding, and ExportAllSheets disabled, ensuring commas and double‑quotes in the data are preserved only when needed.
// Keywords: Aspose.Cells CSV export | export range to CSV .NET | preserve commas in CSV | text qualifier Aspose.Cells | TxtSaveOptions CSV delimiter | C# Aspose.Cells example | ExportArea CSV Aspose
// Common Searches: Aspose.Cells export selected range to CSV | how to keep commas and quotes in CSV export .NET | set text qualifier for CSV in Aspose.Cells | export specific cells as CSV using C# | CSV delimiter options Aspose.Cells
// Developer Intent: Save only the A1:B3 range as a CSV file while retaining commas and quotation marks where they appear in the data.
// Use Cases: Create a CSV excerpt of a product list where descriptions contain commas and quotes, enabling downstream parsers to read the file correctly. | Generate a UTF‑8 encoded CSV report for a worksheet subsection that will be imported into a database expecting quoted fields only when delimiters are present. | Provide a localized CSV export of a data slice (e.g., A1:B3) without quoting every cell, reducing file size and preserving readability.
// AI Prompts: Show how to change the CSV separator to a semicolon while keeping the Normal quote type in Aspose.Cells. | Give an example of exporting multiple non‑contiguous ranges to separate CSV files with Aspose.Cells for .NET. | Explain how to configure a custom text qualifier such as single quotes for CSV export using Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

namespace ExportRangeToCsvDemo
{
    // Demonstrates how to export a defined CellArea (A1:B3) to a CSV file using Aspose.Cells. The example sets TxtSaveOptions with a comma separator, Normal quote type, UTF‑8 encoding, and ExportAllSheets disabled, ensuring commas and double‑quotes in the data are preserved only when needed.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (including delimiters and quotes)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Description");
            cells["A2"].PutValue("Widget");
            // Description contains a comma and quotes, which should be preserved in CSV
            cells["B2"].PutValue("Small, \"lightweight\" widget");
            cells["A3"].PutValue("Gadget");
            cells["B3"].PutValue("Multi-purpose gadget");

            // Define the range to export (A1:B3)
            CellArea exportArea = new CellArea
            {
                StartRow = 0,   // Row 0 (A1)
                EndRow = 2,     // Row 2 (A3)
                StartColumn = 0, // Column 0 (A)
                EndColumn = 1    // Column 1 (B)
            };

            // Configure text save options for CSV export
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',',                 // Use comma as delimiter
                ExportArea = exportArea,         // Export only the defined range
                QuoteType = TxtValueQuoteType.Normal, // Quote only when necessary (preserves delimiters)
                Encoding = Encoding.UTF8,        // Ensure proper encoding
                ExportAllSheets = false          // Export only the active sheet
            };

            // Save the selected range to a CSV file
            workbook.Save("ExportedRange.csv", saveOptions);

            Console.WriteLine("Range exported to CSV successfully.");
        }
    }
}
