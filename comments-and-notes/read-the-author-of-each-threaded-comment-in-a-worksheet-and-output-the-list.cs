// Title: Read and list threaded comment authors from an Excel worksheet using Aspose.Cells in C#
// AI Prompts: Generate C# code that opens a workbook, iterates through all threaded comments, and prints each comment’s author name together with its cell address. | Write a method that returns a distinct collection of ThreadedCommentAuthor objects from an Aspose.Cells workbook. | Demonstrate how to save the workbook after extracting threaded comment author information, including directory‑creation handling.
// Common Searches: Aspose.Cells C# get author names of threaded comments in a worksheet | How to enumerate threaded comment authors in an Excel file with .NET | Retrieve distinct list of comment authors from Aspose.Cells workbook | C# example for reading threaded comments and their authors using Aspose.Cells | Save Excel workbook after processing threaded comments with Aspose.Cells
// Tags: read threaded comment authors Aspose.Cells | enumerate worksheet threaded comments C# | extract unique comment authors Excel .NET | save workbook after comment processing Aspose.Cells | list comment authors per cell Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentAuthorsDemo
{
    // The sample creates a workbook, adds two threaded comment authors, inserts threaded comments into cells, iterates through each comment in the first worksheet, prints the cell address and author name of every threaded comment, and finally saves the workbook to ThreadedCommentAuthorsOutput.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Access the collection of threaded comment authors
                ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;

                // Add two authors
                int authorIdx1 = authors.Add("Alice Johnson", "alice@example.com", "PROV1");
                int authorIdx2 = authors.Add("Bob Smith", "bob@example.com", "PROV2");

                ThreadedCommentAuthor author1 = authors[authorIdx1];
                ThreadedCommentAuthor author2 = authors[authorIdx2];

                // Add threaded comments to different cells using the authors
                sheet.Comments.AddThreadedComment("A1", "First comment by Alice", author1);
                sheet.Comments.AddThreadedComment("A1", "Reply by Bob", author2);
                sheet.Comments.AddThreadedComment("B2", "Another comment by Alice", author1);

                // Iterate through all comments in the worksheet
                Console.WriteLine("Threaded comment authors in the worksheet:");
                foreach (Comment comment in sheet.Comments)
                {
                    // Each comment may contain a collection of threaded comments
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Determine the cell name for the comment using its row/column indices
                    string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);

                    foreach (ThreadedComment tc in threadedComments)
                    {
                        // Output the cell name and author name of each threaded comment
                        Console.WriteLine($"- Cell \"{cellName}\": {tc.Author.Name}");
                    }
                }

                // Save the workbook (optional, just to demonstrate save rule)
                string outputPath = "ThreadedCommentAuthorsOutput.xlsx";

                try
                {
                    // Ensure the directory exists (if a directory part is present)
                    string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
