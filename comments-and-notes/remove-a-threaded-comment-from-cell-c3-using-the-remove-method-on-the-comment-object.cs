// Title: Remove a Threaded Comment from Cell C3 with Aspose.Cells for .NET
// Description: Shows how to add a threaded comment author, place a threaded comment in cell C3, delete it using Worksheet.Comments.RemoveAt, and save the workbook as ThreadedCommentRemoved.xlsx.
// Keywords: Aspose.Cells | .NET | threaded comment | remove comment | Worksheet.Comments.RemoveAt | C3 | Excel comment deletion | programmatic comment removal | Aspose.Cells API | Excel automation
// Common Searches: Aspose.Cells remove threaded comment C3 | How to delete a comment from a specific cell using Aspose.Cells .NET | Worksheet.Comments.RemoveAt example | Delete Excel threaded comments with C# | Remove comments before saving workbook Aspose
// Developer Intent: Programmatically delete the threaded comment attached to cell C3.
// Use Cases: Clean up temporary notes before publishing a workbook to maintain a professional appearance. | Automatically strip user comments after data processing to avoid exposing internal remarks. | Prepare a confidential spreadsheet for external distribution by removing all threaded comments.
// AI Prompts: Provide C# code that removes a threaded comment from cell C3 using Aspose.Cells. | How can I check if a threaded comment exists in C3 before calling RemoveAt with Aspose.Cells? | Explain the steps to delete all threaded comments in a worksheet via the Aspose.Cells API.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentRemoval
{
    // Shows how to add a threaded comment author, place a threaded comment in cell C3, delete it using Worksheet.Comments.RemoveAt, and save the workbook as ThreadedCommentRemoved.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "EXAMPLE_PROVIDER");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell C3
            worksheet.Comments.AddThreadedComment("C3", "Initial threaded comment", author);

            // Remove the comment (including its threaded comments) from cell C3
            worksheet.Comments.RemoveAt("C3");

            // Save the workbook
            workbook.Save("ThreadedCommentRemoved.xlsx");
        }
    }
}
