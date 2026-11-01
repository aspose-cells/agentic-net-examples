// Title: Set a comment shape’s text direction to RightToLeft in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set comment shape text direction to RightToLeft | How to display Arabic text in an Excel comment using Aspose.Cells | Change text direction of a comment shape for bidirectional languages in .NET | Set TextDirection property of a comment shape in an Aspose.Cells workbook | Right-to-left orientation for Excel comments with Aspose.Cells C#
// Tags: comment shape text direction Aspose.Cells | right-to-left comment orientation .NET | bidirectional language support Excel Aspose.Cells | set TextDirectionType RightToLeft C# | Arabic comment rendering Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds an Arabic comment to cell A1, sets the comment's shape TextDirection to RightToLeft for proper bidirectional display, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        var commentIndex = sheet.Comments.Add("A1");
        var comment = sheet.Comments[commentIndex];
        comment.Note = "مثال على نص عربي"; // Sample Arabic text

        // Set the comment's shape text direction to RightToLeft for bidirectional support
        comment.CommentShape.TextDirection = TextDirectionType.RightToLeft;

        // Save the workbook (lifecycle save)
        workbook.Save("output.xlsx");
    }
}
