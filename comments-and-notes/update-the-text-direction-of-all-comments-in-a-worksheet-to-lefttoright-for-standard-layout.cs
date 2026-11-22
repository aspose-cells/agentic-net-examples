// Title: How to set comment text direction to LeftToRight for all comments in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change all worksheet comment text direction to left-to-right | Set left-to-right text direction for Excel comments with Aspose.Cells .NET | Iterate through comments and update TextDirection property in Aspose.Cells | Aspose.Cells comment shape TextDirection left to right example | Programmatically modify Excel comment orientation using Aspose.Cells C#
// Tags: Aspose.Cells CommentShape TextDirection | C# update Excel comment direction | iterate worksheet comments Aspose.Cells | set left-to-right comment layout .NET | modify Excel comment orientation programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates one), accesses the first worksheet, loops through each comment, sets each CommentShape.TextDirection to TextDirectionType.LeftToRight, and saves the workbook as output.xlsx.
class UpdateCommentTextDirection
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if present; otherwise create a new one with a sample comment
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                int commentIndex = ws.Comments.Add("A1");
                Comment sampleComment = ws.Comments[commentIndex];
                sampleComment.Note = "Sample comment";
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Update text direction for each comment in the worksheet
            foreach (Comment comment in worksheet.Comments)
            {
                if (comment.CommentShape != null)
                {
                    comment.CommentShape.TextDirection = TextDirectionType.LeftToRight;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
