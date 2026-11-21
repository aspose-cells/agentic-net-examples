// Title: How to freeze the first three rows of an Excel worksheet using Aspose.Cells FreezePanes in C#
// Common Searches: C# Aspose.Cells freeze first three rows of worksheet | Freeze top rows in Excel file with Aspose.Cells API example | Using FreezePanes to lock rows 1‑3 in a workbook with Aspose.Cells | Aspose.Cells C# code to freeze rows without affecting columns | Create Excel file and freeze header rows using Aspose.Cells in .NET
// Tags: Aspose.Cells FreezePanes API C# | freeze worksheet rows Aspose.Cells | C# Excel row freeze using Aspose.Cells | Aspose.Cells workbook header freeze .NET | Excel FreezePanes method Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first worksheet, and freezes the top three rows by calling FreezePanes with row index 3 and column index 0 while supplying the sheet's total row and column counts. It ensures the output directory exists, saves the file as "FrozenRows.xlsx", and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Freeze the top three rows (row index 3, column index 0)
            // Use the 4‑parameter overload: row, column, totalRows, totalColumns
            sheet.FreezePanes(
                3,                                   // rows to freeze
                0,                                   // columns to freeze
                sheet.Cells.Rows.Count,              // total rows in the sheet
                sheet.Cells.Columns.Count);          // total columns in the sheet

            // Define output file path
            string outputPath = "FrozenRows.xlsx";

            // Ensure the output directory exists (handle cases where the path has no directory part)
            string fullPath = Path.GetFullPath(outputPath);
            string? outputDir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
