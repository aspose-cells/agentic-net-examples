// Title: Add a threaded comment to cell J12, reply programmatically, and save the workbook with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, define two ThreadedCommentAuthor objects, insert an initial threaded comment into cell J12, retrieve the comment thread, add a reply from a second author, and save the file as ThreadedComment_J12.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells threaded comment C# | add reply to threaded comment Aspose.Cells | ThreadedCommentAuthor example | comment thread cell J12 Aspose.Cells | save workbook with comments Aspose.Cells | .NET Excel comment API | Aspose.Cells tutorial
// Common Searches: how to add a threaded comment to a specific cell with Aspose.Cells .NET | reply to an existing threaded comment in Aspose.Cells C# | set up multiple authors for Excel comments using Aspose.Cells | Aspose.Cells example for comment threads | save Excel file with threaded comments Aspose.Cells
// Developer Intent: Create a threaded comment on cell J12, add a reply from another author, and persist the workbook.
// Use Cases: Initialize multiple ThreadedCommentAuthor objects and assign them to comment threads. | Insert a root threaded comment into a designated cell (e.g., J12) via the worksheet Comments API. | Retrieve the ThreadedCommentCollection for a cell and append additional replies before saving. | Generate Excel reports that include collaborative comment threads.
// AI Prompts: Write C# code to add three replies to an existing threaded comment in Aspose.Cells, each using a different author and custom formatting. | Show how to iterate over all threaded comments in a worksheet and output the comment text, author name, and timestamp. | Provide an example that deletes a specific reply from a threaded comment thread using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    // Demonstrates how to create a new workbook, define two ThreadedCommentAuthor objects, insert an initial threaded comment into cell J12, retrieve the comment thread, add a reply from a second author, and save the file as ThreadedComment_J12.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 1. Create authors for the threaded comments
            // -----------------------------------------------------------------
            // Author for the initial comment
            int authorIndex1 = worksheet.Workbook.Worksheets.ThreadedCommentAuthors.Add(
                "Alice",               // Name
                "alice@example.com",   // UserId
                "PROVIDER_1");         // ProviderId
            ThreadedCommentAuthor author1 = worksheet.Workbook.Worksheets.ThreadedCommentAuthors[authorIndex1];

            // Author for the reply comment
            int authorIndex2 = worksheet.Workbook.Worksheets.ThreadedCommentAuthors.Add(
                "Bob",
                "bob@example.com",
                "PROVIDER_2");
            ThreadedCommentAuthor author2 = worksheet.Workbook.Worksheets.ThreadedCommentAuthors[authorIndex2];

            // -----------------------------------------------------------------
            // 2. Add a threaded comment to cell J12
            // -----------------------------------------------------------------
            // This adds the first thread (the root comment) to the specified cell.
            worksheet.Comments.AddThreadedComment("J12", "Initial threaded comment on J12.", author1);

            // -----------------------------------------------------------------
            // 3. Retrieve the threaded comment collection for J12 and add a reply
            // -----------------------------------------------------------------
            ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments("J12");

            // Add a reply to the existing thread
            threadedComments.Add("This is a reply to the initial comment.", author2);

            // -----------------------------------------------------------------
            // 4. Save the workbook
            // -----------------------------------------------------------------
            workbook.Save("ThreadedComment_J12.xlsx");
        }
    }
}
