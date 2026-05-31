using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the comments collection of the worksheet
        CommentCollection comments = worksheet.Comments;

        // Determine the last used row in the worksheet to limit the loop
        int lastRow = worksheet.Cells.MaxDataRow;

        // Column B has index 1 (zero‑based)
        int columnBIndex = 1;

        // Iterate through each row that may contain a threaded comment in column B
        for (int row = 0; row <= lastRow; row++)
        {
            // Retrieve threaded comments for the cell at (row, column B)
            ThreadedCommentCollection threadedComments = comments.GetThreadedComments(row, columnBIndex);

            // If there are any threaded comments, list their authors
            if (threadedComments != null && threadedComments.Count > 0)
            {
                foreach (ThreadedComment tc in threadedComments)
                {
                    // Output the author name of each threaded comment
                    Console.WriteLine($"Cell B{row + 1}: Author = {tc.Author.Name}");
                }
            }
        }
    }
}