// Title: Save a Workbook with Frozen Rows to XLSB Using Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, freeze the top row with FreezePanes, configure XlsbSaveOptions for binary compression, and save the file as a compact XLSB document.
// Keywords: Aspose.Cells FreezePanes C# | save workbook as XLSB | XlsbSaveOptions compression | binary Excel file Aspose.Cells | freeze top row XLSB | Aspose.Cells .NET export | compressed XLSB example
// Common Searches: Aspose.Cells freeze first row and save as XLSB | C# XlsbSaveOptions compression level | How to use FreezePanes with Aspose.Cells | Save workbook to binary XLSB format | Aspose.Cells example for frozen panes and compression
// Developer Intent: Create a workbook, freeze the header row, and export it as a compressed XLSB file using Aspose.Cells for .NET.
// Use Cases: Generate a report with a frozen header row and distribute it as a small‑size XLSB file. | Export large data sets with frozen panes while minimizing attachment size. | Programmatically build a template that includes frozen rows and store it in binary format for faster loading.
// AI Prompts: Write C# code with Aspose.Cells to freeze the first two rows and save the workbook as an XLSB file using maximum compression. | Explain how to adjust the CompressionType in XlsbSaveOptions and when to set ExportAllColumnIndexes. | Provide step‑by‑step instructions to apply FreezePanes with custom row and column parameters before saving to XLSB.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, freeze the top row with FreezePanes, configure XlsbSaveOptions for binary compression, and save the file as a compact XLSB document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Row 2");
            sheet.Cells["B3"].PutValue(200);

            // Freeze the first row (rows above row index 1 are frozen)
            // Use the 4‑parameter overload: totalRows, totalColumns, rowsToFreeze, columnsToFreeze
            sheet.FreezePanes(0, 0, 1, 0);

            // Create XLSB save options (binary compression)
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                // Optional: specify compression level (default is Level6)
                CompressionType = OoxmlCompressionType.Level6,
                ExportAllColumnIndexes = true
            };

            // Save the workbook as an XLSB file using the save options
            workbook.Save("FrozenRows.xlsb", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
