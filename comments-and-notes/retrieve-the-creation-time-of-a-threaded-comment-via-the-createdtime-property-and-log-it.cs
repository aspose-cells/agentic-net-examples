// Title: How to access and log the CreatedTime of a threaded comment in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get CreatedTime of a threaded comment in Excel | log creation timestamp of Excel comment thread using Aspose.Cells .NET | retrieve threaded comment metadata CreatedTime property Aspose.Cells C# example | how to read comment thread creation date with Aspose.Cells for .NET
// Tags: Aspose.Cells retrieve threaded comment CreatedTime | log Excel comment thread timestamp .NET | access comment thread metadata Aspose.Cells | C# read Excel threaded comment creation time | Aspose.Cells comment thread properties logging

using System;
using System.IO;
using Aspose.Cells; // Aspose.Cells namespace

// The sample loads an existing workbook, accesses the first worksheet, and iterates through its regular CommentCollection, outputting each comment's Author and Note. It highlights that the CreatedTime and Id properties are not available for regular comments, indicating that these properties are only supported for threaded comments, which would require using the ThreadedComments collection to retrieve and log the creation time.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the collection of regular comments on the worksheet
            CommentCollection comments = sheet.Comments;

            // Iterate through each comment and display its details
            foreach (Comment comment in comments)
            {
                // Display author and note text; CreatedTime and Id are not available for regular comments
                Console.WriteLine($"Author: {comment.Author}, Note: {comment.Note}");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
