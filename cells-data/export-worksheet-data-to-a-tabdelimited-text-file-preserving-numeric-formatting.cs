// Title: Export a worksheet to a tab‑delimited TSV file while preserving numeric display formatting with Aspose.Cells (C#)
// Description: Creates a workbook, adds product and price data, applies a two‑decimal numeric style, and uses TxtSaveOptions (TSV format, tab separator, UTF‑8, DisplayStyle, minimal quoting) to save the active worksheet as a .tsv file that retains the formatted numbers.
// Keywords: Aspose.Cells | C# | .NET | TSV export | tab delimited | numeric formatting | TxtSaveOptions | DisplayStyle | Workbook.Save | UTF-8 encoding
// Common Searches: Aspose.Cells export to TSV with original number format | C# save worksheet as tab delimited file preserving decimals | TxtSaveOptions DisplayStyle example | How to keep Excel numeric formatting in a .tsv export | Aspose.Cells tab separator for text export
// Developer Intent: Save a worksheet as a tab‑delimited text file that keeps the cell’s displayed numeric format.
// Use Cases: Generate price‑list reports for downstream systems that require tab‑delimited input while maintaining two‑decimal precision. | Provide data feeds to legacy applications that accept TSV files, ensuring numbers appear exactly as shown in Excel. | Automate batch conversion of multiple worksheets to TSV with consistent numeric display across all files.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to a .tsv file, preserving numeric display formatting and using UTF‑8 encoding. | Show how to configure TxtSaveOptions for a tab‑delimited export with DisplayStyle format strategy and minimal quoting. | Explain the steps to apply a two‑decimal numeric style to a column and export the workbook so the TSV output reflects that formatting.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExportTabDelimited
{
    // Creates a workbook, adds product and price data, applies a two‑decimal numeric style, and uses TxtSaveOptions (TSV format, tab separator, UTF‑8, DisplayStyle, minimal quoting) to save the active worksheet as a .tsv file that retains the formatted numbers.
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

                // Populate sample data
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");
                cells["A2"].PutValue("Laptop");
                cells["B2"].PutValue(999.99);
                cells["A3"].PutValue("Phone");
                cells["B3"].PutValue(599.99);

                // Apply numeric format to the price column (two decimal places)
                Style priceStyle = workbook.CreateStyle();
                priceStyle.Number = 2; // Display as number with two decimal places

                // Apply the style to the range B2:B3
                Aspose.Cells.Range priceRange = cells.CreateRange("B2:B3");
                priceRange.SetStyle(priceStyle);

                // Configure TxtSaveOptions for tab‑delimited export
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv) // Tsv = tab‑delimited
                {
                    Separator = '\t',                     // Explicitly set tab character as separator
                    FormatStrategy = CellValueFormatStrategy.DisplayStyle, // Preserve cell display formatting
                    Encoding = Encoding.UTF8,             // Use UTF‑8 encoding
                    QuoteType = TxtValueQuoteType.Minimum // Minimal quoting
                };

                // Export the active worksheet to a tab‑delimited text file
                string outputPath = "ExportedData.tsv";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Worksheet exported to '{outputPath}' with tab delimiters and numeric formatting preserved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
