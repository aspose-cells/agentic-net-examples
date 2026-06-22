using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to the header cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "Header description";
        comment.Author = "Admin";
        comment.Font.Name = "Calibri";
        comment.Font.Size = 11;
        comment.IsVisible = true; // make the comment visible by default

        // Freeze the row that contains the comment (row 1) so it stays visible while scrolling
        // Freeze panes at cell A2, freezing 1 row and 0 columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("HeaderCommentWithFreeze.xlsx");
    }
}