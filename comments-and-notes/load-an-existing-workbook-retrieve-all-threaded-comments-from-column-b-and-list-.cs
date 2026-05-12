using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ThreadedCommentsFromColumnB
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Column B has index 1 (zero‑based)
            int targetColumn = 1;

            // Collection to keep track of unique authors (optional)
            HashSet<string> uniqueAuthors = new HashSet<string>();

            // Iterate through all used rows in the worksheet
            int maxRow = worksheet.Cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                // Retrieve threaded comments for the cell at (row, targetColumn)
                ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, targetColumn);

                // If there are no threaded comments, continue to next row
                if (threadedComments == null || threadedComments.Count == 0)
                    continue;

                // Process each threaded comment
                foreach (ThreadedComment comment in threadedComments)
                {
                    // Get the author name
                    string authorName = comment.Author?.Name ?? "Unknown";

                    // Output the author
                    Console.WriteLine($"Cell {CellsHelper.CellIndexToName(row, targetColumn)} - Comment Author: {authorName}");

                    // Store unique author names (optional)
                    uniqueAuthors.Add(authorName);
                }
            }

            // Optionally, list all distinct authors found in column B
            Console.WriteLine("\nDistinct authors in column B:");
            foreach (string author in uniqueAuthors)
            {
                Console.WriteLine(author);
            }

            // Save the workbook (if any modifications were made)
            workbook.Save("output.xlsx");
        }
    }
}