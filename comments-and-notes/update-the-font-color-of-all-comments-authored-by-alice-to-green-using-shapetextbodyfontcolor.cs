// Title: Change the font color of Excel comments authored by Alice to green using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change comment font color for a specific author | How to set Excel comment shape font color to green in .NET | Update comment text color based on author name using Aspose.Cells | C# iterate workbook comments and apply green font to Alice's notes | Programmatically filter Excel comments by author and modify font color
// Tags: set comment shape font color Aspose.Cells | filter comments by author Aspose.Cells | change Excel comment text color .NET | apply green font to comment author | modify comment appearance using Aspose.Cells | commentshape font color manipulation C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;
using System.IO;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, loops through each worksheet and its comments, identifies comments authored by "Alice", sets the comment shape's font color to green, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all comments in the worksheet
                    foreach (Comment comment in sheet.Comments)
                    {
                        // Check if the comment author is Alice
                        if (string.Equals(comment.Author, "Alice", StringComparison.OrdinalIgnoreCase))
                        {
                            // Set the font color of the entire comment to green
                            // This avoids needing TextBody/Paragraph/Portion classes
                            comment.CommentShape.Font.Color = Color.Green;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
