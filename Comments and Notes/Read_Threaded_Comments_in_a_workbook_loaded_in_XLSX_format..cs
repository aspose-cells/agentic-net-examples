using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentsReader
{
    class Program
    {
        static void Main()
        {
            // Load the workbook from an existing XLSX file
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can change the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of comments in the worksheet
            CommentCollection comments = worksheet.Comments;

            // Iterate through each comment in the collection
            foreach (Comment comment in comments)
            {
                // Check if the comment is a threaded comment
                if (comment.IsThreadedComment)
                {
                    // Get the cell address of the comment
                    string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                    Console.WriteLine($"Threaded comments in cell {cellName}:");

                    // Retrieve the threaded comments for this cell
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Iterate through each threaded comment and display its details
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                        Console.WriteLine($"- Author: {authorName}");
                        Console.WriteLine($"  Notes : {tc.Notes}");
                        Console.WriteLine($"  Row   : {tc.Row}, Column: {tc.Column}");
                        Console.WriteLine($"  Created: {tc.CreatedTime}");
                    }

                    Console.WriteLine(); // Blank line for readability
                }
            }

            // Optionally, you can save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}