// Title: Update all worksheet comment text direction to LeftToRight using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample comments, loops through the worksheet's Comments collection, and sets each comment's CommentShape.TextDirection to TextDirectionType.LeftToRight before saving the file as CommentsLeftToRight.xlsx.
// Keywords: Aspose.Cells comment text direction | LeftToRight comment orientation .NET | CommentShape TextDirection property | Excel comment layout Aspose | C# set comment direction
// Common Searches: Aspose.Cells set comment direction left to right | C# change Excel comment text orientation | Iterate worksheet comments Aspose.Cells | CommentShape TextDirection example | How to force left‑to‑right comments in Excel with Aspose
// Developer Intent: Apply LeftToRight text direction to every comment in a worksheet.
// Use Cases: Standardize comment layout for left‑to‑right languages before sharing the workbook. | Maintain consistent comment appearance when converting Excel files to PDF or images. | Automate comment formatting after bulk‑adding notes programmatically.
// AI Prompts: Write C# code with Aspose.Cells that changes comment text direction to RightToLeft for a selected worksheet. | Provide a reusable method that accepts a Workbook and sets all comment shapes to LeftToRight, handling missing shapes safely. | Explain the impact of TextDirectionType on comment rendering across different locale settings.

using System;
using Aspose.Cells;

// Creates a workbook, adds sample comments, loops through the worksheet's Comments collection, and sets each comment's CommentShape.TextDirection to TextDirectionType.LeftToRight before saving the file as CommentsLeftToRight.xlsx.
class UpdateCommentTextDirection
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample comments to demonstrate the change
            int commentIdx1 = worksheet.Comments.Add("A1");
            worksheet.Comments[commentIdx1].Note = "First comment";

            int commentIdx2 = worksheet.Comments.Add("B2");
            worksheet.Comments[commentIdx2].Note = "Second comment";

            // Update the text direction of every comment to LeftToRight
            foreach (Comment comment in worksheet.Comments)
            {
                // The TextDirection property belongs to the comment's shape
                comment.CommentShape.TextDirection = TextDirectionType.LeftToRight;
            }

            // Save the workbook with the updated comment settings
            workbook.Save("CommentsLeftToRight.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
