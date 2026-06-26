using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        // Path to the workbook to be cleaned
        string inputPath = "input.xlsx";
        // Path where the cleaned workbook will be saved
        string outputPath = "cleaned.xlsx";

        // Load the workbook from file
        Workbook workbook = new Workbook(inputPath);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Get the collection of comments for the current worksheet
            CommentCollection comments = worksheet.Comments;

            // Store the positions (row, column) of comments that have empty or whitespace notes
            List<(int Row, int Column)> emptyCommentPositions = new List<(int, int)>();

            // Scan all comments
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];

                // If the comment text is null, empty, or consists only of whitespace, mark it for removal
                if (string.IsNullOrWhiteSpace(comment.Note))
                {
                    emptyCommentPositions.Add((comment.Row, comment.Column));
                }
            }

            // Remove the identified empty comments
            foreach (var pos in emptyCommentPositions)
            {
                comments.RemoveAt(pos.Row, pos.Column);
            }
        }

        // Save the cleaned workbook
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}