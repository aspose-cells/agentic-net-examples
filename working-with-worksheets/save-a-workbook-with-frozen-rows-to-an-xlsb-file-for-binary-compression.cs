// Title: Save a workbook with frozen rows to XLSB using Aspise.Cells .NET compression
// Description: C# example that creates a Workbook, freezes the first row with FreezePanes, configures XlsbSaveOptions for Level 6 compression (or any OoxmlCompressionType), and saves the file as a binary XLSB (FrozenRows.xlsb).
// Keywords: Aspose.Cells freeze rows | C# FreezePanes | XlsbSaveOptions compression | save workbook as XLSB | binary Excel file Aspose | .NET Excel compression | export frozen panes XLSB
// Common Searches: Aspose.Cells freeze first row and save as XLSB | XlsbSaveOptions compression level example | C# freeze panes then export to binary Excel | how to reduce XLSB file size with Aspose.Cells | save workbook with frozen header row .NET
// Developer Intent: Create a workbook, freeze the top row, and export it as a compressed XLSB file using Aspose.Cells for .NET.
// Use Cases: Generate reports where the header row stays visible while delivering a compact XLSB file. | Automate export of large tables with frozen panes to minimize file size and improve load speed. | Prepare reusable templates that keep the first row fixed and use binary format for faster Excel opening.
// AI Prompts: Show how to freeze multiple rows and columns and save to XLSB with maximum compression using Aspose.Cells .NET. | Provide code to open an existing workbook, apply FreezePanes to row 2, and save with OoxmlCompressionType.Level9. | Explain the effect of ExportAllColumnIndexes when saving a frozen‑pane workbook to XLSB.

using System;
using Aspose.Cells;

// C# example that creates a Workbook, freezes the first row with FreezePanes, configures XlsbSaveOptions for Level 6 compression (or any OoxmlCompressionType), and saves the file as a binary XLSB (FrozenRows.xlsb).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data (optional)
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue(2);
            worksheet.Cells["C2"].PutValue(3);

            // Freeze the first row (row index 1, column index 0)
            // The fourth and fifth parameters specify how many rows and columns to freeze
            worksheet.FreezePanes(1, 0, 1, 0);

            // Create XLSB save options for binary compression
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                // Use default compression level (Level6) or set a different level if needed
                CompressionType = OoxmlCompressionType.Level6,
                ExportAllColumnIndexes = true
            };

            // Save the workbook as an XLSB file using the specified options
            workbook.Save("FrozenRows.xlsb", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
