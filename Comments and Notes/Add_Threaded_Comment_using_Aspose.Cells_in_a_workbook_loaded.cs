using System;
using Aspose.Cells;

namespace ThreadedCommentDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author to the workbook
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
                "Demo Author",          // Author name
                "demo_user@example.com", // User ID (email or any identifier)
                "DemoProvider");        // Provider ID
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell C2 (row index 1, column index 2)
            worksheet.Comments.AddThreadedComment(1, 2, "This is a threaded comment added via Aspose.Cells.", author);

            // Optional: retrieve and display the added threaded comment details
            ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(1, 2);
            foreach (ThreadedComment tc in threadedComments)
            {
                Console.WriteLine($"Comment at C2 by {tc.Author.Name}: {tc.Notes}");
            }

            // Save the workbook with the new threaded comment
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}