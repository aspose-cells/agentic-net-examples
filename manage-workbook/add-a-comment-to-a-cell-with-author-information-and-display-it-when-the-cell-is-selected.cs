// Title: Add an author‑attributed comment to a cell and show it on selection with Aspose.Cells for .NET (C#)
// Common Searches: how to add a comment with author to a specific cell using Aspose.Cells C# | make cell comment visible only when the cell is selected in Aspose.Cells for .NET | set comment author and show comment on cell selection programmatically Aspose.Cells | Aspose.Cells example for displaying comment on cell click C#
// Tags: Aspose.Cells cell comment author attribute | comment.IsVisible property Aspose.Cells C# | display comment on cell selection Aspose.Cells | programmatic workbook comment creation Aspose.Cells | C# add comment to worksheet cell Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, adds a comment authored by "John Doe" to cell A1, sets the comment to be visible when the cell is selected, and saves the file as CommentWithAuthor.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a value in cell A1 (optional, just to have something in the cell)
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample Data");

        // Add a comment to cell A1
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];

        // Set the comment text and author information
        comment.Note = "This is a comment with author information.";
        comment.Author = "John Doe";

        // Make the comment visible when the cell is selected
        comment.IsVisible = true;

        // Save the workbook to a file
        workbook.Save("CommentWithAuthor.xlsx");
    }
}
