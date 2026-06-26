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

        // Column B has index 1 (zero‑based)
        int columnIndex = 1;

        // Determine the last used row in the worksheet
        int lastRow = worksheet.Cells.MaxDataRow;

        // Iterate through all rows in column B
        for (int row = 0; row <= lastRow; row++)
        {
            // Retrieve threaded comments for the current cell
            ThreadedCommentCollection threadedComments = comments.GetThreadedComments(row, columnIndex);

            // If there are any threaded comments, list their authors
            if (threadedComments != null && threadedComments.Count > 0)
            {
                foreach (ThreadedComment tc in threadedComments)
                {
                    // Output the cell address and author name
                    Console.WriteLine($"Cell B{row + 1}: Author = {tc.Author.Name}");
                }
            }
        }

        // Save the workbook (no modifications made, but required by lifecycle rules)
        workbook.Save("output.xlsx");
    }
}