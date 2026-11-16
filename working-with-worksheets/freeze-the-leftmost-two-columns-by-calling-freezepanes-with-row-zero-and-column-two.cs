// Title: Freeze the first two columns of an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to freeze the leftmost two columns in Excel | How to use FreezePanes to lock first two columns in a .NET workbook | Aspose.Cells FreezePanes parameters for freezing columns only | Saving an Excel file after freezing columns using Aspose.Cells C# | Freeze columns without rows in Aspose.Cells worksheet
// Tags: Aspose.Cells FreezePanes left columns | C# freeze first two columns Excel | Aspose.Cells worksheet freeze panes usage | save workbook with frozen columns Aspose.Cells | Excel column freeze Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first worksheet, calls FreezePanes(0,2,0,2) to freeze the leftmost two columns, ensures the output folder exists, and saves the file as FrozenColumns.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the leftmost two columns (row index = 0, column index = 2)
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(0, 2, 0, 2);

            // Define output file path
            string outputPath = "FrozenColumns.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
