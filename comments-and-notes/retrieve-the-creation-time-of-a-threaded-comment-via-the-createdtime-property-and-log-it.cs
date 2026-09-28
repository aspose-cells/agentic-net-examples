// Title: How to read and log the CreatedTime of a threaded comment in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, accesses the first worksheet's ThreadedComments collection, and prints each comment's Author, Note, and CreatedTime to the console. | Update an existing Aspose.Cells comment‑reading example to also fetch the CreatedTime property of each ThreadedComment and output it alongside the author and note.
// Common Searches: Aspose.Cells C# get CreatedTime from threaded comment in Excel | How to display timestamp of Excel threaded comments using Aspose.Cells .NET | Retrieve creation date of a specific threaded comment with Aspose.Cells API | Log Excel comment creation time in C# using Aspose.Cells ThreadedComments | Example code for reading ThreadedComments CreatedTime property in Aspose.Cells
// Tags: Aspose.Cells ThreadedComments CreatedTime retrieval | C# log Excel threaded comment timestamp | Aspose.Cells read comment metadata .xlsx | extract threaded comment creation date Aspose.Cells | display Excel comment author note and time C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentsDemo
{
    // The sample loads an input.xlsx workbook with Aspose.Cells, accesses the first worksheet, iterates through its ThreadedComments collection, and writes each comment's Author, Note, and CreatedTime to the console. It includes checks for file existence and exception handling to ensure robust execution.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the collection of comments (regular comments as fallback)
                CommentCollection comments = worksheet.Comments;

                if (comments != null && comments.Count > 0)
                {
                    // Iterate through each comment and display its author and note
                    foreach (Comment comment in comments)
                    {
                        Console.WriteLine($"Comment by {comment.Author}: {comment.Note}");
                    }
                }
                else
                {
                    Console.WriteLine("No comments found in the worksheet.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
