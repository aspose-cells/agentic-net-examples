using System;
using Aspose.Cells;

class ReadThreadedComments
{
    static void Main()
    {
        // Load the workbook from an existing XLSX file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // load

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Access the comment collection of the current worksheet
            CommentCollection comments = sheet.Comments;

            // Loop through each comment in the collection
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];

                // Process only threaded comments
                if (comment.IsThreadedComment)
                {
                    // Determine the cell address of the comment
                    string cellAddress = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                    Console.WriteLine($"Threaded comments in cell {cellAddress}:");

                    // Retrieve the threaded comment collection for this comment
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Output each threaded comment's text and author
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                        Console.WriteLine($" - {tc.Notes} (by {authorName})");
                    }
                }
            }
        }

        // No modifications are made, so no need to save the workbook
    }
}