// Title: Update all worksheet comment text direction to LeftToRight with Aspose.Cells (C#)
// Description: C# example that creates a workbook, adds comments, loops through each comment in the first worksheet, sets CommentShape.TextDirection to TextDirectionType.LeftToRight, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | Excel comment | text direction | LeftToRight | CommentShape | TextDirectionType | bulk comment update | worksheet automation
// Common Searches: Aspose.Cells set comment text direction left to right | C# change Excel comment orientation with Aspose | update all comment shapes TextDirection in a worksheet | bulk modify comment layout Aspose.Cells .NET
// Developer Intent: Apply LeftToRight text direction to every comment in a worksheet.
// Use Cases: Standardize comment layout for left‑to‑right languages before exporting reports. | Ensure consistent appearance of comments when generating PDFs from Excel files. | Automate bulk comment formatting after programmatically adding notes.
// AI Prompts: Generate C# code using Aspose.Cells to set comment TextDirection to RightToLeft for selected cells. | Explain how to conditionally set comment text direction based on cell language with Aspose.Cells. | Show how to modify comment TextDirection for a specific range in an existing workbook.

using System;
using Aspose.Cells;

namespace UpdateCommentTextDirection
{
    // C# example that creates a workbook, adds comments, loops through each comment in the first worksheet, sets CommentShape.TextDirection to TextDirectionType.LeftToRight, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample comments
                int idx1 = worksheet.Comments.Add("A1");
                worksheet.Comments[idx1].Note = "First comment";

                int idx2 = worksheet.Comments.Add("B2");
                worksheet.Comments[idx2].Note = "Second comment";

                // Update text direction for each comment
                foreach (Comment comment in worksheet.Comments)
                {
                    // Set the text direction of the comment's shape
                    comment.CommentShape.TextDirection = TextDirectionType.LeftToRight;
                }

                // Save the workbook
                string outputPath = "Comments_TextDirection_LeftToRight.xlsx";
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
