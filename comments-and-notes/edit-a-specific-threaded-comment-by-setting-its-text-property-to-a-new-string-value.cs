using System;
using System.IO;
using Aspose.Cells;

class EditThreadedComment
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Cell that contains the threaded comment to edit
            string targetCell = "A1";

            // Retrieve threaded comments for the cell (may be null if none exist)
            ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(targetCell);

            if (threadedComments != null && threadedComments.Count > 0)
            {
                // Edit the first threaded comment
                ThreadedComment comment = threadedComments[0];
                comment.Notes = "Updated comment text set via Aspose.Cells.";
            }
            else
            {
                Console.WriteLine($"No threaded comments found at cell {targetCell}.");
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}