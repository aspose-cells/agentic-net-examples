using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class UpdateCommentTextDirection
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

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

                // Update text direction for each comment
                foreach (Comment comment in worksheet.Comments)
                {
                    if (comment.CommentShape != null)
                    {
                        comment.CommentShape.TextDirection = TextDirectionType.LeftToRight;
                    }
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}