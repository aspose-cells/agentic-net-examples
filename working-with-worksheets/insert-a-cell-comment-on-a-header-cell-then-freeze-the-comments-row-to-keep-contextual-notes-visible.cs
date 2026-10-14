// Title: Add a visible comment to a header cell and freeze the first row using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a comment with author "Admin" and blue font into cell A1, set it to visible, then freeze row 1 in an Aspose.Cells workbook. | Create C# code that writes a header in A1, attaches a styled comment, applies FreezePanes to keep the comment row fixed, and saves the file as CommentAndFreeze.xlsx.
// Common Searches: Aspose.Cells C# add visible comment to header and keep it shown after freezing panes | How to freeze the top row while preserving cell comments in Aspose.Cells .NET | Set comment author and font color for an Excel header using Aspose.Cells C# | Freeze first row without hiding comments in an Aspose.Cells workbook
// Tags: comment styling Aspose.Cells C# | freeze panes top row Aspose.Cells | header cell annotation Aspose.Cells | preserve comment visibility Aspose.Cells | styled cell comment Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The program creates a workbook, writes a header in A1, adds a visible, blue‑font comment authored by Admin, freezes the first row, and saves the file as CommentAndFreeze.xlsx.
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

            // Write header text in cell A1
            Cell headerCell = sheet.Cells["A1"];
            headerCell.PutValue("Product");

            // Add a comment to the header cell (A1)
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This column contains product names.";
            comment.IsVisible = true;               // Show the comment by default
            comment.Author = "Admin";
            comment.Font.Color = Color.Blue;
            comment.Font.Size = 10;

            // Freeze the first row (row index 1) while keeping all columns unfrozen
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("CommentAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
