using System;
using Aspose.Cells;

namespace RetrieveThreadedComments
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing XLSX workbook (uses the provided load rule)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can change the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the comment collection of the worksheet
            CommentCollection comments = worksheet.Comments;

            // Iterate through all comments in the worksheet
            foreach (Comment comment in comments)
            {
                // Check if the comment has threaded comments
                if (comment.IsThreadedComment || comment.ThreadedComments.Count > 0)
                {
                    // Retrieve threaded comments using the comment's ThreadedComments collection
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Output the cell address for context
                    string cellAddress = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                    Console.WriteLine($"Threaded comments for cell {cellAddress}:");

                    // Iterate through each threaded comment and display its details
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                        Console.WriteLine($"- Author: {authorName}, Notes: {tc.Notes}");
                    }

                    Console.WriteLine(); // Blank line for readability
                }
            }

            // Optionally, save the workbook if any modifications were made (uses the provided save rule)
            // workbook.Save("output.xlsx");
        }
    }
}