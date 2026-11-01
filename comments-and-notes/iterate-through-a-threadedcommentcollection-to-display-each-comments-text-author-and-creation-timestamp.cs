// Title: C# example: loop through Excel worksheet comments and print note and author with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to read comments from an Excel worksheet | example code to list comment text and author using Aspose.Cells for .NET | iterate over CommentCollection in Aspose.Cells C# | retrieve Excel comment author and note with Aspose.Cells library | display all comments in first worksheet using Aspose.Cells C#
// Tags: enumerate comment collection Aspose.Cells C# | read comment note Aspose.Cells .NET | get comment author Aspose.Cells Excel | process worksheet comment objects Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file, verifies its existence, accesses the first worksheet, obtains its CommentCollection, and iterates through each Comment to output the Note (comment text) and Author. It also notes that the Comment class does not expose a creation timestamp property.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of regular comments
            CommentCollection comments = worksheet.Comments;

            // Iterate through each comment
            foreach (Comment comment in comments)
            {
                Console.WriteLine("Comment Text: " + comment.Note);
                Console.WriteLine("Author: " + comment.Author);
                // The Comment class does not expose a CreatedTime property; this line is omitted.
                Console.WriteLine(new string('-', 30));
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
