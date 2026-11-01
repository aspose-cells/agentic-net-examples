// Title: How to set specific column widths and freeze the first three columns in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that sets column A to 20 characters, column B to 30, column C to 25, then freezes columns A‑C while leaving rows unfrozen. | Provide a C# snippet that calls Worksheet.FreezePanes to anchor the first three columns while keeping rows scrollable. | Generate a complete Aspose.Cells example that creates a workbook, adjusts column widths in characters, freezes the first three columns, and saves the file as .xlsx.
// Common Searches: Aspose.Cells set column width in characters C# example | Aspose.Cells example for freezing three columns in a worksheet | How to apply custom column widths and freeze panes using Aspose.Cells for C# | C# Aspose.Cells FreezePanes method column only no rows | Set column A width 20 and lock columns A to C with Aspose.Cells
// Tags: set column width Aspose.Cells C# | freeze panes columns Aspose.Cells | custom column sizing Aspose.Cells workbook | Aspose.Cells FreezePanes method usage | Excel .xlsx column width adjustment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, sets column A, B, and C widths to 20, 30, and 25 characters respectively, freezes the first three columns using FreezePanes, and saves the result as CustomColumnWidthsAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set specific column widths (in characters)
            // Column A (index 0) width = 20 characters
            sheet.Cells.SetColumnWidth(0, 20);
            // Column B (index 1) width = 30 characters
            sheet.Cells.SetColumnWidth(1, 30);
            // Column C (index 2) width = 25 characters
            sheet.Cells.SetColumnWidth(2, 25);

            // Freeze the first three columns (A, B, C)
            // FreezePanes(row, column, totalRows, totalColumns)
            // Row = 0 (no rows frozen), Column = 3 (freeze columns before index 3)
            sheet.FreezePanes(0, 3, 0, 0);

            // Define output file path
            string outputPath = "CustomColumnWidthsAndFreeze.xlsx";

            // Ensure the directory exists (handle possible null from GetDirectoryName)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
