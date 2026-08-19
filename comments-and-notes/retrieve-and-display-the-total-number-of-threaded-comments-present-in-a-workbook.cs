// Title: Count all threaded comments in an Aspose.Cells workbook using C#
// Description: Demonstrates how to create a workbook, add a threaded comment author, insert threaded comments, and iterate through every worksheet to sum the ThreadedComments collection. The total threaded comment count is printed to the console and the workbook can be saved.
// Keywords: Aspose.Cells | C# | threaded comments | comment count | workbook API | ThreadedCommentAuthors | IsThreadedComment | ThreadedComments collection | Excel automation | sample code
// Common Searches: Aspose.Cells count threaded comments C# | total threaded comments in Excel workbook | how to sum threaded comments Aspose.Cells | C# iterate worksheets Aspose.Cells comments | retrieve threaded comment count programmatically
// Developer Intent: Calculate the total number of threaded comments across all worksheets in an Aspose.Cells workbook.
// Use Cases: Generate a report of comment activity for quality‑control audits. | Enforce a maximum comment limit before publishing a workbook. | Log workbook comment metrics for analytics or compliance.
// AI Prompts: Create a reusable C# method that returns the total threaded comment count for any Aspose.Cells Workbook. | Write C# code that outputs each worksheet name with its individual threaded comment total using Aspose.Cells. | Produce a unit test that validates the threaded comment counting logic when multiple worksheets contain threaded comments.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCount
{
    // Demonstrates how to create a workbook, add a threaded comment author, insert threaded comments, and iterate through every worksheet to sum the ThreadedComments collection. The total threaded comment count is printed to the console and the workbook can be saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "JD");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add threaded comments to a few cells
            worksheet.Comments.AddThreadedComment("B2", "First comment", author);
            worksheet.Comments.AddThreadedComment("B2", "Reply to first", author);
            worksheet.Comments.AddThreadedComment("C3", "Another comment", author);

            // Calculate total number of threaded comments in the workbook
            int totalThreadedComments = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                CommentCollection comments = ws.Comments;
                foreach (Comment comment in comments)
                {
                    // Only threaded comments have a ThreadedComments collection
                    if (comment.IsThreadedComment)
                    {
                        totalThreadedComments += comment.ThreadedComments.Count;
                    }
                }
            }

            // Display the result
            Console.WriteLine($"Total threaded comments in the workbook: {totalThreadedComments}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ThreadedCommentsCount.xlsx");
        }
    }
}
