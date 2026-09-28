// Title: Add an author‑attributed comment to a cell and make it appear on selection using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a comment into cell A1, assign "John Doe" as the author, set the comment to be visible when the cell is selected, and save the workbook with Aspose.Cells in C#. | Update existing Aspose.Cells code to attach author information to a cell comment and enable automatic display of the comment when the cell gains focus.
// Common Searches: Aspose.Cells C# add comment with author and show on cell click | how to make Excel comment visible on selection using Aspose.Cells .NET | set comment author property in Aspose.Cells workbook C# example | display cell comment automatically when selected Aspose.Cells | C# Aspose.Cells comment.IsVisible true usage
// Tags: add cell comment with author Aspose.Cells | comment visibility on selection .NET | Aspose.Cells comment author property | display Excel comment automatically C# | set comment.IsVisible Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, writes "Sample Data" to cell A1, adds a comment authored by "John Doe", configures the comment to be visible when the cell is selected, and saves the file as CommentExample.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a value in cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample Data");

        // Add a comment to cell A1
        // The comment includes author information and will be displayed when the cell is selected
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is a comment with author information.";
        comment.Author = "John Doe";
        comment.IsVisible = true; // Show comment when the cell is selected

        // Save the workbook (lifecycle rule)
        workbook.Save("CommentExample.xlsx");
    }
}
