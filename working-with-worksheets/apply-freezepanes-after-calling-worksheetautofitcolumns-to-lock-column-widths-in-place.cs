// Title: Freeze the first column after auto‑fitting columns with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that auto‑fits all worksheet columns and then freezes column A using Aspose.Cells. | Show the correct order of calling Worksheet.AutoFitColumns followed by Worksheet.FreezePanes to keep column widths fixed in an Excel file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# freeze column A after AutoFitColumns example | how to lock column widths after calling AutoFitColumns in Aspose.Cells .NET | C# code to freeze first column programmatically using Aspose.Cells | freeze panes after adjusting column sizes with Aspose.Cells for .NET
// Tags: auto-fit columns then freeze panes Aspose.Cells | freeze first column C# Aspose.Cells | preserve column widths after AutoFitColumns .NET | Worksheet.FreezePanes usage Aspose.Cells | Excel column width lock Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, auto‑fits all columns on the first worksheet, freezes column A with FreezePanes(0,1,0,0) to lock the column widths, saves the file as FrozenColumns.xlsx, and includes basic error handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var sheet = workbook.Worksheets[0];

                // Auto‑fit all columns based on their content
                sheet.AutoFitColumns();

                // Freeze the first column (column A)
                // FreezePanes(row, column, totalRows, totalColumns)
                sheet.FreezePanes(0, 1, 0, 0);

                // Save the workbook to a file
                string outputPath = "FrozenColumns.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
