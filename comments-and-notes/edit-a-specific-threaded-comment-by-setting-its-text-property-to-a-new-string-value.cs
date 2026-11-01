// Title: How to update the text of a specific threaded comment in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change existing comment text in Excel worksheet | update comment note property in Aspose.Cells workbook programmatically | modify threaded comment at a given index using Aspose.Cells .NET | set new text for an Excel comment with Aspose.Cells API | replace comment content in the first worksheet using Aspose.Cells C#
// Tags: Aspose.Cells update comment text C# | modify Excel comment note property | set threaded comment content Aspose.Cells | change worksheet comment programmatically | Aspose.Cells comment editing example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an Excel file with Aspose.Cells, accesses the first worksheet, checks for a comment at a specified index, updates the comment's Note property with new text, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Index of the comment to update (adjust as needed)
            int commentIndex = 0;

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the requested comment exists
                if (worksheet.Comments == null || worksheet.Comments.Count <= commentIndex)
                {
                    Console.WriteLine($"Comment at index {commentIndex} does not exist.");
                    return;
                }

                // Retrieve the comment and update its text
                Comment comment = worksheet.Comments[commentIndex];
                comment.Note = "This is the updated comment text.";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
