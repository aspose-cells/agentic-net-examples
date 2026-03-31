using System;
using Aspose.Cells;

class RetrieveThreadedComments
{
    static void Main()
    {
        // Load the workbook (replace with a MemoryStream if the file is already in memory)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the comment collection of the current worksheet
            CommentCollection comments = sheet.Comments;

            // Loop through all comments in the collection
            foreach (Comment comment in comments)
            {
                // Process only threaded comments
                if (comment.IsThreadedComment)
                {
                    // Get the cell address for the comment (e.g., "B2")
                    string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                    Console.WriteLine($"Threaded comments in worksheet \"{sheet.Name}\" cell {cellName}:");

                    // Retrieve the threaded comment collection for this cell
                    ThreadedCommentCollection threadedComments = comments.GetThreadedComments(comment.Row, comment.Column);

                    // Output each threaded comment preserving its order (hierarchy is implicit by order)
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        Console.WriteLine($"  Author   : {tc.Author.Name}");
                        Console.WriteLine($"  Notes    : {tc.Notes}");
                        Console.WriteLine($"  Created  : {tc.CreatedTime}");
                        Console.WriteLine($"  Position : Row {tc.Row}, Column {tc.Column}");
                        Console.WriteLine();
                    }
                }
            }
        }

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx");
    }
}