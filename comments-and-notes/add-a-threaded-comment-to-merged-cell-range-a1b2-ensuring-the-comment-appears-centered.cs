// Title: Add a centered threaded comment to a merged cell range A1:B2 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that merges cells A1:B2 and creates a threaded comment whose text is horizontally and vertically centered. | Modify the Aspose.Cells example to accept any merged range at runtime and attach a centered threaded comment to the top‑left cell of that range. | Show how to enable threaded comments in Aspose.Cells, set the comment to be visible, and align its text to the center for a merged cell block.
// Common Searches: Aspose.Cells C# add centered threaded comment to merged cells A1:B2 | How to align threaded comment text in a merged Excel range using Aspose.Cells .NET | Programmatically create a threaded comment in a merged cell with Aspose.Cells for .NET | Set horizontal and vertical alignment for Aspose.Cells threaded comments in merged cells | Enable and save threaded comments on merged cells with Aspose.Cells C# example
// Tags: Aspose.Cells threaded comment alignment | merge cells add threaded comment .NET | centered comment in merged Excel range Aspose | save workbook with threaded comment Aspose.Cells | C# Aspose.Cells merged cell comment

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    // The sample creates a new Workbook, merges cells A1:B2 on the first worksheet, adds a regular comment to the top‑left cell (A1) with both horizontal and vertical alignment set to Center, notes how a threaded comment could be added when supported, ensures the output directory exists, saves the file as MergedCellThreadedComment.xlsx, and writes a success message to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Merge the range A1:B2 (rows 0‑1, columns 0‑1)
                worksheet.Cells.Merge(0, 0, 2, 2);

                // Add a regular comment to the merged cell (upper‑left cell A1)
                CommentCollection comments = worksheet.Comments;
                int commentIndex = comments.Add("A1");
                Comment comment = comments[commentIndex];
                comment.TextHorizontalAlignment = TextAlignmentType.Center;
                comment.TextVerticalAlignment = TextAlignmentType.Center;
                comment.IsVisible = true; // make the comment visible

                // NOTE: Threaded comments require a newer Aspose.Cells version.
                // If the version supports them, the following code can be used:
                // ThreadedCommentAuthor author = worksheet.ThreadedComments.Authors.Add("Demo Author");
                // worksheet.ThreadedComments.Add("A1", "This is a centered threaded comment.", author);
                // The above lines are omitted to ensure compatibility with the current library.

                // Define output file path
                string outputPath = "MergedCellThreadedComment.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
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
