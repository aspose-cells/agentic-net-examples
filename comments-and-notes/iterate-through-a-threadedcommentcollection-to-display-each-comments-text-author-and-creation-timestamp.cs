// Title: C# example: enumerate worksheet comments and display note, author, and placeholder timestamp with Aspose.Cells
// Common Searches: c# list all comments in an Excel worksheet using Aspose.Cells | how to get comment author and text from a workbook with Aspose.Cells .NET | display Excel comment details in console using Aspose.Cells C# | retrieve comment creation time placeholder Aspose.Cells comment enumeration | aspose.cells read comment note and author from first sheet
// Tags: Aspose.Cells enumerate worksheet comments | C# extract comment note and author | Excel comment placeholder timestamp Aspose | load workbook and read comments Aspose.Cells | comment collection iteration C# Aspose

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, accesses the first worksheet, iterates through its CommentCollection, and writes each comment's note, author, and a placeholder for the creation time to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of regular comments
            CommentCollection comments = worksheet.Comments;

            // Iterate through each comment and display its details
            foreach (Comment comment in comments)
            {
                // Note: In Aspose.Cells, the comment text is accessed via the Note property
                Console.WriteLine("Comment: " + comment.Note);
                Console.WriteLine("Author: " + comment.Author);
                // CreatedTime property is not available; display placeholder
                Console.WriteLine("Created: N/A");
                Console.WriteLine(new string('-', 30));
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
