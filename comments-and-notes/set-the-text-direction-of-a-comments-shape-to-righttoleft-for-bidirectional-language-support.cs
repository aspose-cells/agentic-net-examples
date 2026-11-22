// Title: Set a comment shape's text direction to RightToLeft with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set comment shape text direction to RightToLeft | How to display Arabic text in Excel comment using Aspose.Cells | C# example for bidirectional language support in Excel comments with Aspose.Cells | Change text direction of a comment's shape in Aspose.Cells workbook | Right-to-left comment rendering in .NET Excel using Aspose.Cells
// Tags: comment shape TextDirection property Aspose.Cells | RightToLeft text orientation Excel comment .NET | Arabic language comment shape Aspose.Cells | bidirectional comment support Aspose.Cells C# | set comment shape direction Excel file

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds an Arabic comment to cell A1, sets the comment's shape TextDirection to RightToLeft for proper bidirectional rendering, and saves the file as Comment_RTL.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        var commentIndex = sheet.Comments.Add("A1");
        var comment = sheet.Comments[commentIndex];
        comment.Note = "مثال على نص عربي";

        // Set the comment's shape text direction to RightToLeft
        comment.CommentShape.TextDirection = TextDirectionType.RightToLeft;

        // Save the workbook (lifecycle rule)
        workbook.Save("Comment_RTL.xlsx");
    }
}
