using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ChangeCommentFontColor
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all comments in the worksheet
            foreach (Comment comment in worksheet.Comments)
            {
                // Column G has zero‑based index 6
                if (comment.Column == 6)
                {
                    // Get the shape associated with the comment
                    CommentShape shape = comment.CommentShape;

                    // Change the font color of the comment text to blue
                    shape.Font.Color = Color.Blue;

                    // Ensure the comment is visible (optional)
                    comment.IsVisible = true;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}