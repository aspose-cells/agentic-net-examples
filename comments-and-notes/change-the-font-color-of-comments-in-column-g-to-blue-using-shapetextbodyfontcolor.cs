// Title: Set comment font color to blue in column G using Aspose.Cells C#
// Description: Shows how to create a workbook, add visible comments to cells in column G, access each comment's CommentShape, change its Font.Color to blue, and save the result as CommentsColumnG_BlueFont.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells comment font color | CommentShape Font.Color C# | Excel comment formatting Aspose | change comment text color .NET | blue comment font Aspose.Cells | column G comment styling
// Common Searches: Aspose.Cells set comment text color blue | How to change comment font color in a specific column using .NET | CommentShape Font.Color example Aspose.Cells | Programmatically format Excel comments with Aspose | C# change comment color column G Aspose
// Developer Intent: Apply a blue font color to every comment located in column G of an Excel worksheet.
// Use Cases: Make comments in a report column stand out with a consistent blue font. | Standardize the appearance of bulk‑added comments in a generated spreadsheet. | Visually differentiate notes in a specific column for easier review.
// AI Prompts: Write C# code that changes comment font color to red in column H using Aspose.Cells. | Explain how CommentShape.TextBody.Font.Color can be used to set comment text color in Aspose.Cells. | Provide an example of conditional comment formatting based on cell values with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add visible comments to cells in column G, access each comment's CommentShape, change its Font.Color to blue, and save the result as CommentsColumnG_BlueFont.xlsx with Aspose.Cells for .NET.
class ChangeCommentFontColorInColumnG
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample comments to cells in column G (index 6)
            for (int row = 0; row < 5; row++)
            {
                // Add a comment to the cell at (row, column G)
                int commentIndex = worksheet.Comments.Add(row, 6);
                Comment comment = worksheet.Comments[commentIndex];
                comment.Note = $"Comment at G{row + 1}";
                comment.IsVisible = true; // make it visible for verification
            }

            // Iterate through all comments in the worksheet
            foreach (Comment comment in worksheet.Comments)
            {
                // Process only comments that are in column G (index 6)
                if (comment.Column == 6)
                {
                    // Get the shape associated with the comment
                    CommentShape shape = comment.CommentShape;

                    // Set the font color of the comment text to blue
                    shape.Font.Color = Color.Blue;
                }
            }

            // Save the workbook
            workbook.Save("CommentsColumnG_BlueFont.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
