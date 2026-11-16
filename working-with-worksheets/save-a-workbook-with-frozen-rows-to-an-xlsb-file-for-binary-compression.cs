// Title: Freeze the first row of a worksheet and save it as a compressed XLSB file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells freeze first row and export to XLSB | How to use FreezePanes and SaveFormat.Xlsb in Aspose.Cells | Saving a workbook with frozen panes to binary XLSB using Aspose.Cells .NET | Create XLSB file with frozen rows in Aspose.Cells C# example | Aspose.Cells ensure output directory before saving XLSB workbook
// Tags: freeze panes Aspose.Cells C# | save workbook as XLSB Aspose.Cells | binary Excel compression Aspose.Cells | output directory handling Aspose.Cells | worksheet FreezePanes method example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, adds data, freezes the first row with FreezePanes, ensures the target folder exists, and saves the file as a compressed binary XLSB using SaveFormat.Xlsb.
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

            // Add some sample data (optional)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["A3"].PutValue("Data 2");

            // Freeze the first row (freeze panes at row 1, column 0)
            // Parameters: totalRows, totalColumns, rows, columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Define output file path
            string outputPath = "FrozenRowsWorkbook.xlsb";

            // Ensure the directory exists (handle case where outputPath has no directory part)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook as an XLSB file (binary format with compression)
            workbook.Save(outputPath, SaveFormat.Xlsb);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
