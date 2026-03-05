using System;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        // Path to the existing XLSX workbook containing threaded comments
        string inputPath = "input.xlsx";

        // Load the workbook (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the collection of comments on the worksheet
        CommentCollection comments = worksheet.Comments;

        // Iterate through each comment in the collection
        foreach (Comment comment in comments)
        {
            // Retrieve threaded comments for the current comment's cell
            ThreadedCommentCollection threadedComments = comments.GetThreadedComments(comment.Row, comment.Column);

            // Iterate through each threaded comment and display its CreatedTime
            foreach (ThreadedComment tc in threadedComments)
            {
                // Convert row/column to cell name for readability
                string cellName = CellsHelper.CellIndexToName(tc.Row, tc.Column);
                Console.WriteLine($"Threaded comment in cell {cellName} was created at: {tc.CreatedTime}");
            }
        }
    }
}