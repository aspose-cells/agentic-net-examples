// Title: C# – Change Comment Font Color to Blue in Column G with Aspose.Cells Shape.TextBody.Font.Color
// Description: A concise Aspose.Cells for .NET example that loads an Excel workbook, finds all comments in column G, accesses each comment's CommentShape, iterates the TextBody FontSettingCollection, sets Font.Color to blue, optionally makes the comment visible, and saves the file.
// Keywords: Aspose.Cells C# comment color | Shape.TextBody.Font.Color | CommentShape font color | Excel comment blue column G | Aspose.Cells API example | C# Excel comment formatting | GitHub Aspose.Cells sample | programmatic comment styling | .NET Excel automation
// Common Searches: Aspose.Cells set comment text color C# | change font color of Excel comments using Shape.TextBody | blue comment text column G Aspose.Cells | how to format comment appearance with Aspose.Cells .NET | sample code for comment color in Aspose.Cells
// Developer Intent: Programmatically set the font color of every comment located in column G to blue using Aspose.Cells for .NET.
// Use Cases: Visually differentiate comments in a specific column for reporting dashboards. | Enforce a consistent comment color scheme before distributing workbooks to stakeholders. | Automate Excel generation pipelines where comment readability must meet brand guidelines.
// AI Prompts: Generate C# code that changes comment font color to red in column G using Aspose.Cells. | Show how to set both font size and bold style for comments in column G with Shape.TextBody. | Explain how to apply conditional colors to comments based on their row index in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// A concise Aspose.Cells for .NET example that loads an Excel workbook, finds all comments in column G, accesses each comment's CommentShape, iterates the TextBody FontSettingCollection, sets Font.Color to blue, optionally makes the comment visible, and saves the file.
class ChangeCommentFontColor
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Iterate through all comments in the worksheet
        foreach (Comment comment in worksheet.Comments)
        {
            // Check if the comment is in column G (zero‑based index 6)
            if (comment.Column == 6)
            {
                // Get the shape that represents the comment
                CommentShape commentShape = comment.CommentShape;

                // Access the TextBody collection of the shape
                FontSettingCollection textBody = commentShape.TextBody;

                // Apply blue color to all font settings within the TextBody
                foreach (FontSetting fontSetting in textBody)
                {
                    fontSetting.Font.Color = Color.Blue;
                }

                // Optionally make the comment visible
                comment.IsVisible = true;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
