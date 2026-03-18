using System;
using Aspose.Cells;

namespace AsposeCellsCommentDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook (must be an .xlsx file)
            string inputPath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a new comment to cell B2 using the CommentCollection.Add method
            // This returns the index of the newly created comment
            int commentIndex = worksheet.Comments.Add("B2");

            // Retrieve the comment object from the collection
            Comment comment = worksheet.Comments[commentIndex];

            // Set comment properties
            comment.Author = "AsposeDemo";
            comment.Note = "This is a newly added comment.";

            // Save the workbook with the added comment (lifecycle rule: save)
            workbook.Save("output.xlsx");

            Console.WriteLine("Comment added to cell B2 and workbook saved as output.xlsx");
        }
    }
}