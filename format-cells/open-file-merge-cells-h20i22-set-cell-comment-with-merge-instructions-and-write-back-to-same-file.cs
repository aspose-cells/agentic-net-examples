// Title: Merge cells H20:I22, insert a comment, and overwrite the original workbook with Aspose.Cells for .NET
// AI Prompts: Load an existing Excel file, merge the range H20:I22, place a note on cell H20, and save the changes back to the same file using Aspose.Cells in C#. | Using Aspose.Cells for .NET, combine cells H20 through I22, attach a comment to the top‑left cell, and persist the workbook over the original document.
// Common Searches: Aspose.Cells C# merge range H20 I22 and add a note to the merged cell | how to overwrite the original Excel file after merging cells with Aspose.Cells | insert comment into merged cells using Aspose.Cells .NET example | C# code to merge specific cells and save over same workbook with Aspose.Cells | Aspose.Cells merge cells and preserve existing file path
// Tags: merge range H20 I22 Aspose.Cells .NET | add note to merged cell C# | save workbook over original file Aspose.Cells | modify existing Excel workbook Aspose.Cells | cell comment insertion Aspose.Cells C#

using Aspose.Cells;
using System;

// The program loads 'input.xlsx' with Aspose.Cells, merges cells H20 through I22, adds a comment "Merged cells H20:I22" to cell H20, and saves the workbook back to the same file.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells H20:I22 (rows 20-22, columns H and I)
        // Row and column indices are zero‑based: H = 7, 20 = 19
        sheet.Cells.Merge(19, 7, 3, 2); // 3 rows, 2 columns

        // Add a comment to the top‑left cell of the merged range (H20)
        int commentIndex = sheet.Comments.Add("H20");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "Merged cells H20:I22";

        // Save the workbook back to the same file
        workbook.Save(filePath);
    }
}
