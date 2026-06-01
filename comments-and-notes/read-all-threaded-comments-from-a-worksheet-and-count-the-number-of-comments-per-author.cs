using System;
using System.Collections.Generic;
using Aspose.Cells;

class ThreadedCommentAuthorCounter
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("ThreadedCommentsDemo.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Dictionary to hold comment count per author name
        Dictionary<string, int> authorCommentCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Iterate through all comments in the worksheet
        foreach (Comment comment in worksheet.Comments)
        {
            // Each comment may contain a collection of threaded comments
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;

            // Count each threaded comment by its author
            foreach (ThreadedComment tc in threadedComments)
            {
                string authorName = tc.Author?.Name ?? "Unknown";

                if (authorCommentCount.ContainsKey(authorName))
                    authorCommentCount[authorName]++;
                else
                    authorCommentCount[authorName] = 1;
            }
        }

        // Output the results
        Console.WriteLine("Threaded comment count per author:");
        foreach (var kvp in authorCommentCount)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }

        // Optionally, save the workbook if any modifications were made
        // workbook.Save("ThreadedCommentsDemo_Output.xlsx");
    }
}