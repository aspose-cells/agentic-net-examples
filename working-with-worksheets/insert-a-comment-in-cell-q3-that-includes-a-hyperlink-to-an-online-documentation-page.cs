using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell Q3
        int commentIndex = worksheet.Comments.Add("Q3");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Author = "Automation";
        comment.Note = "See the online documentation for details.";

        // Add a hyperlink to the same cell (Q3)
        int hyperlinkIndex = worksheet.Hyperlinks.Add("Q3", 1, 1, "https://docs.aspose.com/cells/net/");
        // Set the display text of the hyperlink
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Aspose.Cells Documentation";

        // Save the workbook
        workbook.Save("CommentWithHyperlink.xlsx");
    }
}