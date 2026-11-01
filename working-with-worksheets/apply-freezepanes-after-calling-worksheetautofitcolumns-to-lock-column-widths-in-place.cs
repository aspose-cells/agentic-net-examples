// Title: How to AutoFit columns then freeze the top row in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that calls Worksheet.AutoFitColumns and then Worksheet.FreezePanes to auto‑size columns and lock the first row in an Aspose.Cells workbook. | Show a .NET example that freezes the header row after auto‑fitting all columns with Aspose.Cells.
// Common Searches: C# Aspose.Cells auto fit columns and then freeze first row example | Freeze panes after AutoFitColumns in Aspose.Cells .NET workbook | How to keep column widths fixed when freezing panes with Aspose.Cells for C#
// Tags: auto‑fit columns then freeze panes Aspose.Cells | Worksheet.AutoFitColumns with FreezePanes C# | lock column widths after autofit Aspose.Cells | freeze top row after column auto‑size .NET | Aspose.Cells workbook column width locking

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, populates cells, auto‑fits all columns, freezes the first row to preserve column widths, and saves the file as Output.xlsx using Aspose.Cells for .NET.
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

            // Sample data to demonstrate AutoFit
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("This is a long piece of text that will cause the column to expand");
            sheet.Cells["B2"].PutValue(12345);

            // Auto‑fit all columns based on the content
            sheet.AutoFitColumns();

            // Freeze the first row (row index 1, column index 0) and keep column widths locked
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 1, 0);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
