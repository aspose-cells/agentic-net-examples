// Title: Change all Excel comment text directions to LeftToRight in a worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set all comment text direction to left-to-right in a worksheet | How to loop through worksheet comments and change text direction with Aspose.Cells | Bulk update Excel comment layout to LeftToRight using .NET API | Aspose.Cells comment shape TextDirection property example in C# | Programmatically change Excel comment direction to LTR with Aspose.Cells
// Tags: comment shape TextDirection Aspose.Cells | enumerate worksheet comment collection C# | apply LeftToRight direction to Excel comments | bulk comment layout modification .NET | using TextDirectionType LeftToRight in code | update comment orientation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, iterates over each comment, and sets the CommentShape.TextDirection property to TextDirectionType.LeftToRight before saving the updated file.
class UpdateCommentTextDirection
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all comments and set text direction
            foreach (Comment comment in sheet.Comments)
            {
                // Ensure the comment has an associated shape
                if (comment.CommentShape != null)
                {
                    comment.CommentShape.TextDirection = TextDirectionType.LeftToRight;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
