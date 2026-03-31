using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Simulate the action of invoking a context menu on cell C3 and adding a comment
        // Add a comment to cell C3 using the CommentCollection.Add(string) method
        int commentIndex = worksheet.Comments.Add("C3");

        // Retrieve the newly added comment object
        Comment comment = worksheet.Comments[commentIndex];

        // Set comment properties (text, author, visibility, etc.)
        comment.Note = "This comment was added programmatically.";
        comment.Author = "AutomationScript";
        comment.IsVisible = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}