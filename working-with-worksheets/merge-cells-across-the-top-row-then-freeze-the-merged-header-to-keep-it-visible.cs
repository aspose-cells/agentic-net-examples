// Title: How to merge the top row into a single header cell and freeze it using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that merges cells A1 through D1, assigns a header label, and freezes the first worksheet row so the header stays visible while scrolling. | Generate a .NET example that creates an Excel file, combines the top‑row cells into one header cell, sets its text, and applies FreezePanes to lock the header row using Aspose.Cells.
// Common Searches: Aspose.Cells C# merge A1:D1 and keep header visible while scrolling | freeze panes after merging top row header in Aspose.Cells .NET | C# example for merging first row cells and freezing the row with Aspose.Cells | how to create a frozen merged header in Excel using Aspose.Cells for .NET
// Tags: merge top row cells Aspose.Cells | freeze first worksheet row Aspose.Cells | Aspose.Cells merged header example | C# create frozen header Excel | Excel workbook header freeze Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, merges cells A1‑D1 into a single header cell, sets the text "Header", freezes the first row so the merged header remains visible during scrolling, and saves the file as MergedHeader.xlsx.
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

            // Merge cells A1 to D1 (top row)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            sheet.Cells.Merge(0, 0, 1, 4);

            // Set a value for the merged header cell
            sheet.Cells[0, 0].PutValue("Header");

            // Freeze the top row so the merged header stays visible while scrolling
            // Overload: FreezePanes(row, column, totalRows, totalColumns)
            // Freeze 1 row above row index 1 (i.e., the first row)
            sheet.FreezePanes(1, 0, 1, 0);

            // Define output file path
            string outputPath = "MergedHeader.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
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
