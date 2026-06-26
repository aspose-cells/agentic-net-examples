using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCommentColorDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (lifecycle rule: load)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the comments collection of the current worksheet
                CommentCollection comments = sheet.Comments;

                // Loop through each comment in the collection
                for (int i = 0; i < comments.Count; i++)
                {
                    Comment comment = comments[i];

                    // Determine the font color based on the comment author
                    // Example logic: "Alice" -> Blue, "Bob" -> Green, others -> Red
                    if (comment.Author != null && comment.Author.Equals("Alice", StringComparison.OrdinalIgnoreCase))
                    {
                        comment.Font.Color = Color.Blue;
                    }
                    else if (comment.Author != null && comment.Author.Equals("Bob", StringComparison.OrdinalIgnoreCase))
                    {
                        comment.Font.Color = Color.Green;
                    }
                    else
                    {
                        comment.Font.Color = Color.Red;
                    }
                }
            }

            // Save the modified workbook to a new file (lifecycle rule: save)
            workbook.Save("output.xlsx");
        }
    }
}