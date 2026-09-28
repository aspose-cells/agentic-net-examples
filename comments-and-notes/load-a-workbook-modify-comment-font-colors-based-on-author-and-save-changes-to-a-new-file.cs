// Title: Set comment font color per author in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, loops through every worksheet and each comment, assigns Font.Color based on the comment’s Author, and saves the result to a new file. | Create a helper method that returns a System.Drawing.Color for a given comment author and apply it to all comments in a workbook via Aspose.Cells.
// Common Searches: c# Aspose.Cells change comment text color according to author | apply unique colors to Excel cell comments per user with Aspose.Cells | loop through all worksheets and update comment colors for each author in .NET | export workbook after modifying comment fonts using Aspose.Cells | sample code mapping comment author to System.Drawing.Color in Aspose.Cells
// Tags: Aspose.Cells set comment font color by author | C# iterate worksheet comments Aspose.Cells | Aspose.Cells author-based comment color mapping | save modified workbook Aspose.Cells C# | Excel comment color customization Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// Loads input.xlsx, iterates each worksheet and its comments, determines a font color for each comment based on its Author via a helper method, applies the color, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Access the comments collection of the current worksheet
            CommentCollection comments = worksheet.Comments;

            // Loop through each comment in the collection
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];

                // Determine the font color based on the comment's author
                Color authorColor = GetColorForAuthor(comment.Author);

                // Apply the determined color to the comment's font
                comment.Font.Color = authorColor;
            }
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }

    // Helper method to map authors to specific colors
    static Color GetColorForAuthor(string author)
    {
        // Example mapping – customize as needed
        if (author.Equals("Alice", StringComparison.OrdinalIgnoreCase))
            return Color.Blue;
        if (author.Equals("Bob", StringComparison.OrdinalIgnoreCase))
            return Color.Green;
        if (author.Equals("Charlie", StringComparison.OrdinalIgnoreCase))
            return Color.Orange;

        // Default color for any other author
        return Color.Red;
    }
}
