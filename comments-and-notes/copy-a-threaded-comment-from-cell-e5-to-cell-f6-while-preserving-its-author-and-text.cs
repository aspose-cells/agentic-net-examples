// Title: Copy a threaded comment from cell E5 to cell F6 while preserving author and text with Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies a threaded comment from a source cell to a target cell, retaining the original Author and Note using Aspose.Cells. | Show how to duplicate an Excel comment programmatically by adding a new comment at a different location and copying its properties with Aspose.Cells in .NET. | Provide a method that moves a comment from one worksheet coordinate to another without losing metadata in a C# application.
// Common Searches: Aspose.Cells copy comment from one cell to another C# | preserve comment author when duplicating Excel comment using Aspose.Cells .NET | how to transfer threaded comment properties between cells in Aspose.Cells | C# example for copying Excel cell comment with author and text | Aspose.Cells duplicate comment to different worksheet location
// Tags: copy threaded comment Aspose.Cells | preserve comment author .NET | duplicate Excel comment C# | Aspose.Cells comment property transfer | move comment between cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCopy
{
    // The example creates a workbook, adds a threaded comment to cell E5, retrieves its Author and Note, adds a new comment to cell F6, copies the original comment's properties to the new comment, and saves the workbook as ThreadedCommentCopy.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Add a regular comment to source cell E5 (row 4, column 4)
                // ------------------------------------------------------------
                int sourceRow = 4; // zero‑based index for row 5
                int sourceCol = 4; // zero‑based index for column E
                string authorName = "John Doe";
                string commentText = "Original comment text";

                // Add the comment (Add returns the comment index)
                worksheet.Comments.Add(sourceRow, sourceCol);
                // Retrieve the newly added comment
                Comment srcComment = worksheet.Comments[sourceRow, sourceCol];
                srcComment.Author = authorName;
                srcComment.Note = commentText;

                // ------------------------------------------------------------
                // 2. Retrieve the comment from the source cell
                // ------------------------------------------------------------
                Comment retrievedComment = worksheet.Comments[sourceRow, sourceCol];
                if (retrievedComment == null)
                {
                    Console.WriteLine("Source comment not found.");
                    return;
                }

                // ------------------------------------------------------------
                // 3. Copy the comment to destination cell F6 (row 5, column 5)
                // ------------------------------------------------------------
                int destRow = 5; // zero‑based index for row 6
                int destCol = 5; // zero‑based index for column F

                // Add a comment at the destination cell
                worksheet.Comments.Add(destRow, destCol);
                // Retrieve the destination comment and copy properties
                Comment destComment = worksheet.Comments[destRow, destCol];
                destComment.Author = retrievedComment.Author;
                destComment.Note = retrievedComment.Note;

                // ------------------------------------------------------------
                // 4. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "ThreadedCommentCopy.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
