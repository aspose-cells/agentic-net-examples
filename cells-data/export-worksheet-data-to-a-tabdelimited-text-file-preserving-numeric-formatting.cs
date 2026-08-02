// Title: Export Worksheet to Tab‑Delimited TSV with Currency Formatting using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply a currency number format to cells, and save the active sheet as a UTF‑8 tab‑delimited TSV file. The example uses TxtSaveOptions with SaveFormat.Tsv, a tab separator, and CellValueFormatStrategy.DisplayStyle to retain the displayed numeric formatting.
// Keywords: Aspose.Cells | C# | export TSV | tab delimited text | currency format | CellValueFormatStrategy | TxtSaveOptions | SaveFormat.Tsv | preserve numeric formatting | UTF-8 encoding | ExportAllSheets | Excel to TSV
// Common Searches: Aspose.Cells export to TSV with formatting | C# save worksheet as tab delimited file | preserve currency format when exporting Excel to text | TxtSaveOptions DisplayStyle example | how to export only active sheet using Aspose.Cells
// Developer Intent: Save a worksheet as a tab‑delimited file while keeping the displayed numeric (currency) format.
// Use Cases: Generate a price‑list TSV for an e‑commerce catalog where prices retain currency symbols. | Provide a tab‑delimited feed to a legacy system that requires formatted numeric values. | Automate daily export of financial reports to TSV, exporting only the active sheet with UTF‑8 encoding.
// AI Prompts: Write C# code with Aspose.Cells to export the active worksheet to a .tsv file, preserving currency formatting. | Show how to configure TxtSaveOptions for UTF‑8, tab separator, DisplayStyle format strategy, and ExportAllSheets = false.

using System;
using System.Text;
using Aspose.Cells;

namespace ExportTabDelimited
{
    // Demonstrates how to create a workbook, apply a currency number format to cells, and save the active sheet as a UTF‑8 tab‑delimited TSV file. The example uses TxtSaveOptions with SaveFormat.Tsv, a tab separator, and CellValueFormatStrategy.DisplayStyle to retain the displayed numeric formatting.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate cells with sample data
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");
                cells["A2"].PutValue("Laptop");
                cells["B2"].PutValue(999.99);
                cells["A3"].PutValue("Phone");
                cells["B3"].PutValue(599.99);

                // Apply a numeric format to the price column (e.g., currency)
                Style priceStyle = workbook.CreateStyle();
                priceStyle.Number = 2; // Currency format

                // Set the style for the range B2:B3 using Aspose.Cells.Range
                Aspose.Cells.Range priceRange = cells.CreateRange("B2:B3");
                priceRange.SetStyle(priceStyle);

                // Configure text save options for tab‑delimited output
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv)
                {
                    Separator = '\t',                     // Tab delimiter
                    Encoding = Encoding.UTF8,             // UTF‑8 encoding
                    FormatStrategy = CellValueFormatStrategy.DisplayStyle, // Preserve formatted numeric values
                    ExportAllSheets = false               // Export only the active sheet
                };

                // Save the worksheet as a tab‑delimited text file
                workbook.Save("ExportedData.tsv", saveOptions);

                Console.WriteLine("Worksheet exported to tab‑delimited file with numeric formatting preserved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
