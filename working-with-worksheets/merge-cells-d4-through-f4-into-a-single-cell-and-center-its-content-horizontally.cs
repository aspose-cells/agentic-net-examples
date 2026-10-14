// Title: How to merge cells D4 to F4 and horizontally center the text using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that merges cells D4 through F4, inserts a header string, applies centered alignment, and saves the workbook as MergedCells.xlsx. | Generate a .NET example that uses Aspose.Cells to combine a three‑column range into one cell, set its text, align it to the middle horizontally, and export the file.
// Common Searches: Aspose.Cells example for merging D4 to F4 and centering the header in C# | C# Aspose.Cells how to align text horizontally in a merged Excel cell | Merge three columns into one cell with Aspose.Cells and save as XLSX | Center text in a merged range using Aspose.Cells for .NET
// Tags: merge cell range D4-F4 Aspose.Cells | set cell value after merging Aspose.Cells | apply style to merged cell Aspose.Cells | export workbook to XLSX Aspose.Cells

using Aspose.Cells;
using System;

// // This program creates a new workbook, merges cells D4 through F4 on the first worksheet, assigns a header value, centers the text horizontally, and saves the result as MergedCells.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells D4 through F4 (row index 3, column index 3, 1 row, 3 columns)
        sheet.Cells.Merge(3, 3, 1, 3);

        // Set a value for the merged cell
        sheet.Cells["D4"].PutValue("Merged Header");

        // Retrieve the style of the merged cell
        Style style = sheet.Cells["D4"].GetStyle();

        // Center the content horizontally
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the updated style back to the cell
        sheet.Cells["D4"].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("MergedCells.xlsx");
    }
}
