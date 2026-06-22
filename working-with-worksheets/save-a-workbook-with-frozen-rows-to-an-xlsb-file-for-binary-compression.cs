using System;
using System.IO;
using Aspose.Cells;

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

            // Add sample data
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue(2);
            worksheet.Cells["C2"].PutValue(3);

            // Freeze the first row (row index 1, column index 0)
            // FreezePanes(totalRows, totalColumns, row, column)
            worksheet.FreezePanes(1, 0, 1, 0);

            // Configure XLSB save options
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                ExportAllColumnIndexes = true // default, kept for clarity
            };

            string outputPath = "FrozenRows.xlsb";

            // Save the workbook as an XLSB file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any runtime errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}