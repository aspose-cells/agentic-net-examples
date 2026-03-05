using System;
using Aspose.Cells;

class EditComment
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the collection of comments
        CommentCollection comments = sheet.Comments;

        if (comments != null && comments.Count > 0)
        {
            // Retrieve the first comment
            Comment comment = comments[0];

            // Update the comment's text
            comment.Note = "This is the updated comment text.";
        }
        else
        {
            Console.WriteLine("No comments found in the worksheet.");
        }

        // Save the modified workbook
        workbook.Save("OutputWorkbook.xlsx");
    }
}