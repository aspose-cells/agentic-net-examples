// Title: How to change Excel comment font colors by author using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set comment font color based on comment author | Iterate through worksheet comments and apply conditional font colors with Aspose.Cells | Change Excel comment text color programmatically for specific authors using Aspose.Cells | Aspose.Cells example to color comments red for Alice and blue for Bob | Save modified workbook after updating comment font colors in C#
// Tags: comment font color author Aspose.Cells | worksheet comment iteration C# | conditional comment color Excel .NET | programmatic comment text styling Aspose.Cells | output workbook after comment modifications

using Aspose.Cells;
using System.Drawing;

// The example loads an existing workbook, iterates through each worksheet and its comments, assigns a font color (red for author Alice, blue for author Bob, black for others), updates the comment font color, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all comments in the current worksheet
            foreach (Comment comment in sheet.Comments)
            {
                // Determine the font color based on the comment's author
                Color fontColor;
                switch (comment.Author)
                {
                    case "Alice":
                        fontColor = Color.Red;
                        break;
                    case "Bob":
                        fontColor = Color.Blue;
                        break;
                    default:
                        fontColor = Color.Black;
                        break;
                }

                // Apply the determined color to the comment's font
                comment.Font.Color = fontColor;
            }
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
