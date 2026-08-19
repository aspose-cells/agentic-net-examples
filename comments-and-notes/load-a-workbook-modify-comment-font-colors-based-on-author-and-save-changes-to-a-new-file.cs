// Title: Set Excel Comment Font Color by Author with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, walks through every worksheet, reads each comment, selects a font color based on the comment's author, applies the color via the Comment.Font property, and saves the updated file as a new workbook.
// Keywords: Aspose.Cells | C# | Excel comment font color | comment author color mapping | modify comment style programmatically | Workbook.Save | Comment.Font | iterate worksheet comments
// Common Searches: Aspose.Cells change comment font color by author | C# set Excel comment color based on author | how to loop through comments with Aspose.Cells | map comment authors to colors in .NET Excel | save workbook after updating comment styles Aspose
// Developer Intent: Update each Excel comment's font color according to its author and write the changes to a new file.
// Use Cases: Visually differentiate reviewer notes in a shared spreadsheet by assigning a unique color to each author. | Create an audit trail that highlights comments per contributor before distributing the workbook to stakeholders. | Automate branding rules that require specific authors' comments to appear in designated colors.
// AI Prompts: Generate C# code using Aspose.Cells that reads a dictionary of author‑color pairs and applies the corresponding font color to every comment. | Show how to change the background fill of comments instead of the font color, based on the comment author. | Explain how to modify additional font attributes (bold, italic, size) for comments while also setting the color with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Loads an Excel workbook, walks through every worksheet, reads each comment, selects a font color based on the comment's author, applies the color via the Comment.Font property, and saves the updated file as a new workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (load rule)
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the comments collection of the worksheet
            var comments = sheet.Comments;

            // Process each comment
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];

                // Determine the font color based on the comment author
                Color authorColor = GetColorForAuthor(comment.Author);

                // Modify the comment's font color (using Comment.Font property)
                comment.Font.Color = authorColor;
            }
        }

        // Save the modified workbook to a new file (save rule)
        string outputFile = "output.xlsx";
        workbook.Save(outputFile);
    }

    // Helper method to map authors to specific colors
    static Color GetColorForAuthor(string author)
    {
        // Example mapping: Alice -> Red, Bob -> Green, others -> Blue
        if (string.Equals(author, "Alice", StringComparison.OrdinalIgnoreCase))
            return Color.Red;
        if (string.Equals(author, "Bob", StringComparison.OrdinalIgnoreCase))
            return Color.Green;
        return Color.Blue;
    }
}
