using System;
using Aspose.Cells;

class RetrieveThreadedComments
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Access the comment collection of the current worksheet
            CommentCollection comments = sheet.Comments;

            // Loop through all comments in the collection
            foreach (Comment comment in comments)
            {
                // Determine if the comment has any threaded comments
                // (IsThreadedComment indicates the comment itself is a threaded comment,
                //  but threaded replies are stored in the ThreadedComments collection.)
                if (comment.IsThreadedComment || comment.ThreadedComments.Count > 0)
                {
                    // Retrieve threaded comments for the cell using row/column indices
                    ThreadedCommentCollection threaded = comments.GetThreadedComments(comment.Row, comment.Column);

                    // Output each threaded comment's text and author
                    foreach (ThreadedComment tc in threaded)
                    {
                        Console.WriteLine($"Cell {CellsHelper.CellIndexToName(comment.Row, comment.Column)}: \"{tc.Notes}\" by {tc.Author.Name}");
                    }
                }
            }
        }
    }
}