using System;
using Aspose.Cells;

namespace ThreadedCommentTimestampDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Access the collection of comments on the current worksheet
                CommentCollection comments = worksheet.Comments;

                // Loop through each comment (each comment is attached to a specific cell)
                for (int i = 0; i < comments.Count; i++)
                {
                    Comment comment = comments[i];

                    // Determine the cell address of the comment (e.g., "A1")
                    string cellAddress = CellsHelper.CellIndexToName(comment.Row, comment.Column);

                    // Retrieve the threaded comments associated with this comment
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Loop through each threaded comment and output its creation timestamp
                    for (int j = 0; j < threadedComments.Count; j++)
                    {
                        ThreadedComment tc = threadedComments[j];
                        Console.WriteLine($"Worksheet: {worksheet.Name}, Cell: {cellAddress}, Threaded Comment #{j + 1}, Created Time: {tc.CreatedTime}");
                    }
                }
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}