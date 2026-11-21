// Title: How to merge the top row cells and freeze the header row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge A1:D1 and keep header visible with FreezePanes | C# code to create a merged header row and freeze the first row in Excel using Aspose.Cells | Freeze top row after merging cells in an Aspose.Cells .NET workbook | Set merged cell value and apply FreezePanes on the first worksheet with Aspose.Cells | Programmatically create a merged header and freeze panes in Aspose.Cells for .NET
// Tags: merge cells top row Aspose.Cells C# | freeze panes first row Aspose.Cells | set merged cell value Aspose.Cells | create merged header worksheet Aspose.Cells | Aspose.Cells workbook save Excel

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, merges cells A1 through D1 on the first worksheet, writes "Header" into the merged cell, freezes the top row so the header stays visible while scrolling, and saves the file as MergedHeader.xlsx.
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

            // Merge cells across the top row (A1 to D1)
            // Parameters: startRow, startColumn, totalRows, totalColumns
            sheet.Cells.Merge(0, 0, 1, 4);

            // Set header text in the merged cell
            sheet.Cells["A1"].PutValue("Header");

            // Freeze the top row so the merged header stays visible while scrolling
            // Freeze rows above row index 1 (i.e., freeze row 1)
            sheet.FreezePanes(1, 0, 0, 0);

            // Define output file path
            string outputPath = "MergedHeader.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
