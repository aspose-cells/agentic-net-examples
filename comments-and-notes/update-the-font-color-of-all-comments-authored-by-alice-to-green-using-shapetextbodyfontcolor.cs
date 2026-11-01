// Title: Change the font color of Excel comments authored by Alice to green using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells change comment font color for author Alice | How to set green font for specific comment author in an Excel workbook using Aspose.Cells | Iterate Excel comments and modify font color based on author with Aspose.Cells .NET | Set comment shape font color to green for a given author using Aspose.Cells C# | Update comment text color for a particular author in a workbook with Aspose.Cells
// Tags: Aspose.Cells comment shape font color | set comment author font color C# | green comment text Aspose.Cells | modify Excel comment font color by author | C# Aspose.Cells comment formatting

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, loops through each worksheet and its comments, identifies comments authored by "Alice", ensures each comment has a shape, sets the comment shape's font color to green, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all comments in the worksheet
                    foreach (Comment comment in sheet.Comments)
                    {
                        // Check if the comment author is Alice
                        if (comment.Author == "Alice")
                        {
                            // Ensure the comment has an associated shape
                            if (comment.CommentShape != null)
                            {
                                // Set the font color of the comment text to green
                                comment.CommentShape.Font.Color = Color.Green;
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
