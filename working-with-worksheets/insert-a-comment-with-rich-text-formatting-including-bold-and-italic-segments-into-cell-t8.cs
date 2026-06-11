using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell T8 (row 7, column 19)
        int commentIndex = worksheet.Comments.Add("T8");
        Comment comment = worksheet.Comments[commentIndex];

        // Set the comment text with rich formatting using HTML:
        // Bold segment and italic segment
        comment.HtmlNote = "<b>Bold segment</b> and <i>Italic segment</i>";

        // Make the comment visible (optional)
        comment.IsVisible = true;

        // Save the workbook
        workbook.Save("CommentRichText_T8.xlsx");
    }
}