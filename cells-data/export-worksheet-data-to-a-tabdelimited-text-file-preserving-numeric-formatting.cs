// Title: Export a worksheet to a tab‑delimited TSV file while preserving numeric display formatting using Aspose.Cells for C#
// AI Prompts: Generate C# code that uses Aspose.Cells to save a worksheet as a .tsv file with tab separators and keeps the cell's display formatting for numbers. | Show how to apply a two‑decimal number style to a column and export the sheet with TxtSaveOptions so the formatted values appear in the TSV output.
// Common Searches: Aspose.Cells C# export worksheet to TSV preserving number formatting | How to keep two decimal places when saving Excel as tab delimited with Aspose.Cells | TxtSaveOptions DisplayStyle format strategy example for TSV output | Save Excel data as .tsv file with tab separator using Aspose.Cells C# | Apply numeric style to column before exporting to tab‑delimited text in Aspose.Cells
// Tags: Aspose.Cells TSV export with display style | TxtSaveOptions tab separator configuration | apply numeric style to column Aspose.Cells | preserve number formatting in TSV output | C# save worksheet as tab‑delimited text

using System;
using System.Text;
using Aspose.Cells;

// Creates a workbook, fills product and price cells, applies a two‑decimal number style to the price column, configures TxtSaveOptions with a tab separator and DisplayStyle strategy, and saves the active worksheet as a .tsv file while preserving the formatted numeric values.
class ExportTabDelimited
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(999.99);
            worksheet.Cells["A3"].PutValue("Phone");
            worksheet.Cells["B3"].PutValue(599.99);

            // Apply numeric formatting (two decimal places) to the price column
            Style priceStyle = workbook.CreateStyle();
            priceStyle.Number = 2; // format as 0.00
            // Use a range to apply the style to B2:B3
            worksheet.Cells.CreateRange("B2:B3").SetStyle(priceStyle);

            // Configure text save options for tab‑delimited output
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv)
            {
                Separator = '\t', // tab delimiter
                FormatStrategy = CellValueFormatStrategy.DisplayStyle // preserve formatted numeric values
            };

            // Export the active worksheet to a tab‑delimited text file
            workbook.Save("ExportedData.tsv", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
