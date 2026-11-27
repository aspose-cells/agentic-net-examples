// Title: Set Excel comment font color based on author using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change comment font color for a specific author | how to apply different colors to Excel comments by author in .NET | programmatically set Excel comment text color using Aspose.Cells | iterate through worksheet comments and modify font color in C# | save workbook after updating comment colors with Aspose.Cells
// Tags: Aspose.Cells comment font color | C# worksheet comment iteration | author based comment styling Excel | programmatic comment appearance update | save modified workbook Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace CommentFontColorModifier
{
    // The example loads an existing workbook (input.xlsx), iterates over every worksheet and its comments, sets the comment font color to red for author "John", blue for author "Jane", and black for all other authors, then saves the updated workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all comments on the current worksheet
                foreach (Comment comment in sheet.Comments)
                {
                    // Change the font color based on the comment author
                    if (comment.Author.Equals("John", StringComparison.OrdinalIgnoreCase))
                    {
                        comment.Font.Color = Color.Red;
                    }
                    else if (comment.Author.Equals("Jane", StringComparison.OrdinalIgnoreCase))
                    {
                        comment.Font.Color = Color.Blue;
                    }
                    else
                    {
                        // Default color for other authors
                        comment.Font.Color = Color.Black;
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}
