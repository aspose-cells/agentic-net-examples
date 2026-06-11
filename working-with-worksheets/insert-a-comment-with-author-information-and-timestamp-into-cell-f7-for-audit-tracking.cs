using System;
using Aspose.Cells;

class InsertComment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Target cell for the audit comment
        string cellName = "F7";

        // Add a comment to the specified cell
        int commentIndex = worksheet.Comments.Add(cellName);
        Comment comment = worksheet.Comments[commentIndex];

        // Set the author of the comment
        comment.Author = "AuditSystem";

        // Include a timestamp in the comment text
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        comment.Note = $"Audit entry created on {timestamp}.";

        // Make the comment visible (optional)
        comment.IsVisible = true;

        // Save the workbook with the comment
        workbook.Save("AuditComment.xlsx");
    }
}