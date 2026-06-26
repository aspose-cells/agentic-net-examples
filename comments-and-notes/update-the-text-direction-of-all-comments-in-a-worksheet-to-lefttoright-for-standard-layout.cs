using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCommentDirectionDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data to cells
                worksheet.Cells["A1"].PutValue("Cell A1");
                worksheet.Cells["B2"].PutValue("Cell B2");
                worksheet.Cells["C3"].PutValue("Cell C3");

                // Add comments to the cells
                int idx1 = worksheet.Comments.Add("A1");
                worksheet.Comments[idx1].Note = "First comment";

                int idx2 = worksheet.Comments.Add("B2");
                worksheet.Comments[idx2].Note = "Second comment";

                int idx3 = worksheet.Comments.Add("C3");
                worksheet.Comments[idx3].Note = "Third comment";

                // Update text direction of all comments to LeftToRight
                foreach (Comment comment in worksheet.Comments)
                {
                    // The CommentShape holds formatting properties, including text direction
                    comment.CommentShape.TextDirection = TextDirectionType.LeftToRight;
                }

                // Define output file path
                string outputPath = "CommentsWithLeftToRightDirection.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
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