using System;
using System.IO;
using Aspose.Cells;

class ModifyCommentOrientation
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "This is a sample comment.";

        // Set the desired text orientation for the comment
        comment.TextOrientationType = TextOrientationType.TopToBottom; // Options: ClockWise, CounterClockWise, NoRotation, TopToBottom

        // Save the workbook to a memory stream (save rule)
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            // The stream now contains the workbook with the updated comment orientation.
        }
    }
}