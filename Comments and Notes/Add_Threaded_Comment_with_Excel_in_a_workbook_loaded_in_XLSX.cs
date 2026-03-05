using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (you can select any worksheet as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author to the workbook
            // Parameters: author name, user id, provider id
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("Demo Author", "demo_user", "demo_provider");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell C3 using row/column indices (row 2, column 2)
            worksheet.Comments.AddThreadedComment(2, 2, "This is a threaded comment added via Aspose.Cells.", author);

            // Alternatively, you could add a threaded comment using the cell name:
            // worksheet.Comments.AddThreadedComment("C3", "This is a threaded comment added via Aspose.Cells.", author);

            // Save the workbook with the new threaded comment
            workbook.Save("output.xlsx");
        }
    }
}