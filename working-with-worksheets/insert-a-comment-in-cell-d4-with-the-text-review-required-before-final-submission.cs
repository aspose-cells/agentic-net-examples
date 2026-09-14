// Title: Add a comment with the text “Review required before final submission” to cell D4 of the first worksheet using Aspose.Cells for .NET
// AI Prompts: Insert a comment containing 'Review required before final submission' into cell D4 and save the workbook as output.xlsx. | Change the target cell and comment text dynamically by using variables before adding the comment with Aspose.Cells. | After adding a comment, retrieve its index from the worksheet's Comments collection and output it to the console. | Add multiple comments to different cells in the same worksheet and verify each comment's note property.
// Common Searches: how to add a note to cell D4 using Aspose.Cells C# | Aspose.Cells set comment text for a specific cell in .NET | retrieve comment index after adding a comment with Aspose.Cells | save Excel file after inserting comments with Aspose.Cells for .NET
// Tags: worksheet.comments.add Aspose.Cells | set comment.note property C# | add comment to cell D4 Aspose.Cells | retrieve comment index Aspose.Cells | save workbook to xlsx Aspose.Cells

using Aspose.Cells;

// // Creates a new workbook, adds a comment with the text 'Review required before final submission' to cell D4 (row 4, column 4) of the first worksheet, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one if needed)
        var workbook = new Workbook();

        // Get the first worksheet
        var worksheet = workbook.Worksheets[0];

        // Add a comment to cell D4 (row index 3, column index 3)
        // The Add method returns the index of the new comment in the Comments collection
        int commentIndex = worksheet.Comments.Add(3, 3);
        var comment = worksheet.Comments[commentIndex];
        comment.Note = "Review required before final submission";

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
