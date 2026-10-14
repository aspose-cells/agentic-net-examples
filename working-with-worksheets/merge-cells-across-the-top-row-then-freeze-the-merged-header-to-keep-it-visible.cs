// Title: How to merge the top row cells and freeze the header row in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new Workbook, merges cells A1 through D1, sets a header text, and freezes the first row with Aspose.Cells. | Show how to make the merge range and freeze row configurable so the same method can be applied to any worksheet using Aspose.Cells for .NET. | Generate a reusable C# function that accepts a Worksheet, a cell range string, a header string, and a freeze row index, then merges the cells, writes the header, applies FreezePanes, and saves the file.
// Common Searches: Aspose.Cells C# merge cells in first row and keep header visible while scrolling | How to freeze top row after merging cells with Aspose.Cells .NET | C# example for merging A1:D1 and applying FreezePanes in Aspose.Cells | Set frozen header for merged cells using Aspose.Cells for .NET
// Tags: Aspose.Cells merge cells top row | Aspose.Cells FreezePanes C# | merged header row Excel Aspose | create workbook with frozen header Aspose.Cells | C# Aspose.Cells set header value merged cells

using Aspose.Cells;
using System;
using System.IO;

// // Demonstrates creating a new workbook, merging cells A1:D1, inserting a header text, and freezing the first row so the merged header stays visible while scrolling, using Aspose.Cells for .NET.
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
            // Parameters: start row, start column, total rows, total columns
            sheet.Cells.Merge(0, 0, 1, 4);

            // Set header text in the merged cell
            sheet.Cells["A1"].PutValue("Header");

            // Freeze the top row so the merged header stays visible while scrolling
            // Freeze row 0 (i.e., rows above row index 1) and no columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Define output file path
            string outputPath = "MergedHeader.xlsx";

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
