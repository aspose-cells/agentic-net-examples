using System;
using System.Drawing;
using Aspose.Cells;

namespace CommentFontColorModifier
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the comments collection of the worksheet
                CommentCollection comments = sheet.Comments;

                // Loop through each comment in the collection
                for (int i = 0; i < comments.Count; i++)
                {
                    Comment comment = comments[i];

                    // Determine the font color based on the comment author
                    // Example logic: Red for "Alice", Blue for "Bob", default Black otherwise
                    if (comment.Author != null && comment.Author.Equals("Alice", StringComparison.OrdinalIgnoreCase))
                    {
                        comment.Font.Color = Color.Red;
                    }
                    else if (comment.Author != null && comment.Author.Equals("Bob", StringComparison.OrdinalIgnoreCase))
                    {
                        comment.Font.Color = Color.Blue;
                    }
                    else
                    {
                        comment.Font.Color = Color.Black;
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");

            // Clean up
            workbook.Dispose();
        }
    }
}