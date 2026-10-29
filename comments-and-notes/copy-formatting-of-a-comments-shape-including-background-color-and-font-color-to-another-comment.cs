// Title: Copy comment shape formatting (background and font color) from one cell comment to another using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy comment shape background color to another cell | How to duplicate Excel comment formatting with Aspose.Cells .NET | Programmatically transfer comment font color between cells using Aspose.Cells | Copy comment style from cell A1 to B2 in C# with Aspose.Cells | Set target comment shape formatting based on source comment in Aspose.Cells
// Tags: comment shape formatting transfer Aspose.Cells | comment background color copy .NET | comment font color replication C# | duplicate Excel comment style Aspose.Cells | apply source comment formatting to target cell | Aspose.Cells comment shape API usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentFormatting
{
    // The example loads an Excel workbook, ensures comments exist in cells A1 and B2, copies the source comment's shape font color (and optionally other font attributes) to the target comment's shape, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // ----- Source comment (cell A1) -----
                Comment sourceComment = sheet.Comments["A1"];
                if (sourceComment == null)
                {
                    // Create a source comment for demonstration
                    int idx = sheet.Comments.Add("A1");
                    sourceComment = sheet.Comments[idx];
                    sourceComment.CommentShape.Text = "Source comment";
                }

                // ----- Target comment (cell B2) -----
                Comment targetComment = sheet.Comments["B2"];
                if (targetComment == null)
                {
                    // Create a target comment for demonstration
                    int idx = sheet.Comments.Add("B2");
                    targetComment = sheet.Comments[idx];
                    targetComment.CommentShape.Text = "Target comment";
                }

                // ----- Copy formatting from source comment shape to target comment shape -----
                // Copy font color
                targetComment.CommentShape.Font.Color = sourceComment.CommentShape.Font.Color;

                // Optionally copy other font attributes
                // targetComment.CommentShape.Font.Name = sourceComment.CommentShape.Font.Name;
                // targetComment.CommentShape.Font.Size = sourceComment.CommentShape.Font.Size;
                // targetComment.CommentShape.Font.IsBold = sourceComment.CommentShape.Font.IsBold;

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
