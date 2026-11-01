// Title: Set the font color of Excel comments in column G to blue using Aspose.Cells Shape.Font in C#
// Common Searches: C# Aspose.Cells change comment text color to blue for a specific column | How to modify Excel comment font color in column G using Shape.Font with Aspose.Cells | Programmatically set blue font for comments in column G using Aspose.Cells .NET | Aspose.Cells iterate comments and apply font color change in C# | Set comment shape font color in Excel file with Aspose.Cells C# example
// Tags: Aspose.Cells comment shape font color C# | Shape.Font.Color usage Aspose.Cells | Excel comment formatting Aspose.Cells | C# iterate worksheet comments Aspose.Cells | apply font color to comment shape

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for Shape classes

// The C# program loads an Excel workbook, loops through all comments on the first worksheet, and for each comment located in column G (index 6) it sets the comment shape's font color to blue via the Shape.Font.Color property, then saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Column G has zero‑based index 6
            const int targetColumn = 6;

            // Iterate through all comments in the worksheet
            foreach (Comment comment in sheet.Comments)
            {
                // Process only comments that are in column G
                if (comment.Column == targetColumn && comment.CommentShape != null)
                {
                    // Change the font color of the entire comment text to blue
                    comment.CommentShape.Font.Color = Color.Blue;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
