// Title: Aspose.Cells for .NET – Set Excel Comment Font Color to Red (C#)
// Description: This C# example creates a workbook, adds a comment to cell A1, accesses its CommentShape, and changes the comment text color to red via TextBody[0].Font.Color before saving as CommentWithRedFont.xlsx.
// Keywords: Aspose.Cells comment font color | C# change Excel comment color | CommentShape TextBody color | set comment text red Aspose | Excel comment styling .NET | Aspose.Cells API TextBody Font.Color
// Common Searches: Aspose.Cells set comment text color | C# change Excel comment font to red | How to format comment shape in Aspose.Cells | Red font for Excel comment using .NET | CommentShape TextBody formatting example
// Developer Intent: Apply a red font color to a worksheet comment programmatically.
// Use Cases: Emphasize critical notes in generated spreadsheets | Flag warnings in automated reports with red comment text | Apply consistent comment styling during batch workbook creation
// AI Prompts: Write C# code with Aspose.Cells to change comment font color to blue. | Provide a loop that sets all worksheet comments to green text in a workbook. | Explain how to use CommentShape.TextBody to modify font size, style, and color of Excel comments.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds a comment to cell A1, accesses its CommentShape, and changes the comment text color to red via TextBody[0].Font.Color before saving as CommentWithRedFont.xlsx.
class ChangeCommentFontColor
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1 and set its text
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This is a sample comment";

            // Access the shape that represents the comment
            CommentShape commentShape = comment.CommentShape;

            // Set the comment text
            commentShape.Text = comment.Note;

            // Change the font color of the comment text to red
            if (commentShape.TextBody.Count > 0)
            {
                commentShape.TextBody[0].Font.Color = Color.Red;
            }

            // Define output file path
            string outputPath = "CommentWithRedFont.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
