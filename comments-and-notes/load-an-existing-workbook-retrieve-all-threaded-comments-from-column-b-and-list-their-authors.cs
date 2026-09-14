// Title: How to read and display authors of threaded comments in column B of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# function that opens an existing .xlsx file with Aspose.Cells, scans column B, and returns a list of author names from all threaded comments. | Generate a .NET code snippet that iterates through each cell in column B, extracts any threaded comments via Aspose.Cells, and prints each comment's author to the console. | Create a reusable method in C# that accepts a worksheet and column index, and outputs the row numbers and authors of all threaded comments found in that column using Aspose.Cells.
// Common Searches: Aspose.Cells C# get threaded comment author names from a specific column | list authors of Excel threaded comments column B using Aspose.Cells | C# iterate over worksheet cells to read threaded comment authors with Aspose.Cells | how to extract threaded comment authors from an .xlsx file in .NET
// Tags: Aspose.Cells threaded comment author extraction | C# read Excel column B comments Aspose.Cells | enumerate worksheet threaded comments .NET | extract comment authors from .xlsx Aspose.Cells

using System;
using Aspose.Cells;

namespace ThreadedCommentsDemo
{
    // Loads input.xlsx, iterates rows in column B, retrieves any threaded comments per cell, and writes each comment's author (or "Unknown Author") to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Column B has index 1 (zero‑based)
            int columnIndex = 1;

            // Get the comments collection of the worksheet
            CommentCollection comments = worksheet.Comments;

            // Determine the range of rows to inspect.
            // Using MaxDataRow ensures we cover rows that contain data;
            // you can adjust this if comments exist beyond data rows.
            int maxRow = worksheet.Cells.MaxDataRow;

            Console.WriteLine("Threaded comment authors in column B:");

            // Iterate through each row in column B
            for (int row = 0; row <= maxRow; row++)
            {
                // Retrieve threaded comments for the current cell (row, column B)
                ThreadedCommentCollection threadedComments = comments.GetThreadedComments(row, columnIndex);

                // If there are any threaded comments, list their authors
                if (threadedComments != null && threadedComments.Count > 0)
                {
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        // Author may be null if not set; guard against it
                        string authorName = tc.Author != null ? tc.Author.Name : "Unknown Author";
                        Console.WriteLine($"Row {row + 1}: {authorName}");
                    }
                }
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}
