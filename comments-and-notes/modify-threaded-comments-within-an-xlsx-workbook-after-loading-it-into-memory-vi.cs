using System;
using Aspose.Cells;

namespace ThreadedCommentModificationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook into memory
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can change the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one threaded comment author.
            // If the author already exists you can retrieve it by index; here we simply add a new one.
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("Demo Author", "demo_user", "demo_provider");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Define the target cell for threaded comment operations
            string targetCell = "B2";

            // Retrieve existing threaded comments for the target cell
            ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(targetCell);

            // If there are existing threaded comments, modify the first one's text
            if (threadedComments != null && threadedComments.Count > 0)
            {
                // Update the Notes property of the first threaded comment
                threadedComments[0].Notes = "Updated comment text (modified by code).";
            }

            // Add a new threaded comment to the same cell
            worksheet.Comments.AddThreadedComment(targetCell, "New threaded comment added programmatically.", author);

            // Save the modified workbook to a new file
            string outputPath = "ModifiedThreadedComments.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Threaded comments have been updated and saved to '{outputPath}'.");
        }
    }
}