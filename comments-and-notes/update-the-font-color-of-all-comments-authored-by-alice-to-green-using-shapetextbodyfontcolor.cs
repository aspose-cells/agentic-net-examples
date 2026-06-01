using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class UpdateCommentFontColor
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of comments on the current worksheet
                CommentCollection comments = sheet.Comments;

                // Loop through each comment
                for (int i = 0; i < comments.Count; i++)
                {
                    Comment comment = comments[i];

                    // Check if the comment author is "Alice"
                    if (string.Equals(comment.Author, "Alice", StringComparison.OrdinalIgnoreCase))
                    {
                        // Get the shape associated with the comment
                        Shape commentShape = comment.CommentShape;

                        // Set the font color of the comment text to green
                        // For comment shapes the Font property controls the text formatting
                        commentShape.Font.Color = Color.Green;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}