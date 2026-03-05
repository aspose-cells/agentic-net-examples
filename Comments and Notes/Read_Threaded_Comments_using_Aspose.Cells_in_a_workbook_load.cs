using System;
using Aspose.Cells;

namespace ThreadedCommentsReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (XLSX format)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can change the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of comments for the worksheet
            CommentCollection comments = worksheet.Comments;

            // Iterate through all cells that have comments
            foreach (Comment comment in comments)
            {
                // Only process threaded comments
                if (comment.IsThreadedComment)
                {
                    // Retrieve the threaded comments for this cell using row/column indices
                    ThreadedCommentCollection threadedComments = comments.GetThreadedComments(comment.Row, comment.Column);

                    // Display cell address
                    string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                    Console.WriteLine($"Threaded comments in cell {cellName}:");

                    // Enumerate each threaded comment
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                        Console.WriteLine($"- [{authorName}] {tc.Notes}");
                    }

                    Console.WriteLine(); // Blank line for readability
                }
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("OutputWorkbook.xlsx");
        }
    }
}