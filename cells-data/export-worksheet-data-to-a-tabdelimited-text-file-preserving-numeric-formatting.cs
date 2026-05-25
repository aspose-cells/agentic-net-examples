using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class ExportWorksheetToTabDelimited
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate cells with sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Laptop");
            cells["B2"].PutValue(999.99);
            cells["A3"].PutValue("Phone");
            cells["B3"].PutValue(599.99);

            // Apply a numeric format to the price column (currency with two decimals)
            Style priceStyle = workbook.CreateStyle();
            priceStyle.Number = 2; // Currency format

            // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range priceRange = worksheet.Cells.CreateRange("B2:B3");
            priceRange.SetStyle(priceStyle);

            // Configure TxtSaveOptions for tab‑delimited export while preserving numeric formatting
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv)
            {
                Separator = '\t',                                 // Tab delimiter
                FormatStrategy = CellValueFormatStrategy.DisplayStyle, // Preserve displayed formatting
                Encoding = Encoding.UTF8,                         // Use UTF‑8 encoding
                QuoteType = TxtValueQuoteType.Normal              // Default quoting behavior
            };

            // Define output file path
            string outputPath = "ExportedData.tsv";

            // Delete existing file if present
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            // Export the active worksheet to a tab‑delimited text file
            workbook.Save(outputPath, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}