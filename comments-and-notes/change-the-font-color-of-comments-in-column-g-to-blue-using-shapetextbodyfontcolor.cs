// Title: Set the font color of Excel comments in column G to blue using Aspose.Cells Shape.TextBody in C#
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, finds all comments in column G, and changes their text color to blue via the comment's Shape.TextBody Font settings. | Write a .NET method that iterates over a worksheet's CommentCollection, filters comments located in column G, and applies Color.Blue to each FontSetting inside the comment's TextBody. | Create a script that loads an existing workbook, updates the font color of comment shapes in column G to blue using Aspose.Cells, and saves the modified file.
// Common Searches: Aspose.Cells C# change comment text color for specific column | How to set comment font color to blue in Excel using Aspose.Cells .NET | Iterate comments in column G and modify shape text color with Aspose.Cells | C# code to update comment shape color in an existing workbook using Aspose.Cells
// Tags: Aspose.Cells comment visual formatting | C# adjust comment font hue | Shape.TextBody Font.Color usage | Excel column G comment update | Aspose.Cells modify comment appearance

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// The program loads an Excel workbook, scans the first worksheet for comments located in column G, accesses each comment's Shape.TextBody, sets the Font.Color of all FontSetting objects to blue, and saves the workbook.
class ChangeCommentFontColor
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the collection of comments on the worksheet
        CommentCollection comments = worksheet.Comments;

        // Iterate through all comments
        for (int i = 0; i < comments.Count; i++)
        {
            Comment comment = comments[i];

            // Check if the comment is in column G (zero‑based index 6)
            if (comment.Column == 6)
            {
                // Get the shape associated with the comment
                Shape commentShape = comment.CommentShape;

                // Access the TextBody collection of the shape
                FontSettingCollection textBody = commentShape.TextBody;

                // Apply blue color to all font settings within the TextBody
                foreach (FontSetting fontSetting in textBody)
                {
                    fontSetting.Font.Color = Color.Blue;
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
