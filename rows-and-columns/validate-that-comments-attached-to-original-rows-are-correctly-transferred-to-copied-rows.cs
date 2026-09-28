// Title: How to verify that cell comments are preserved when copying rows with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to copy a range of rows and duplicate their comments with ShapeCollection.CopyCommentsInRange, then iterate over the destination cells to assert that each comment's Note matches the source. | Create a C# validation routine that, after copying rows via worksheet.Cells.CopyRows, checks that the Comment.Row property of each copied comment equals the new row index and that the comment text remains unchanged.
// Common Searches: Aspose.Cells copy rows and keep cell comments .NET example | validate copied comments after worksheet.Cells.CopyRows in C# | ShapeCollection.CopyCommentsInRange how to use for row duplication | check comment note text after copying rows with Aspose.Cells | C# verify comment row index after copying rows in Aspose.Cells
// Tags: duplicate rows preserving comments Aspose.Cells | ShapeCollection.CopyCommentsInRange usage C# | comment note integrity check Aspose.Cells | verify comment row property after copy | Aspose.Cells comment transfer test

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentCopyValidation
{
    // // Demonstrates adding comments to cells A1 and A2, copying rows 0‑1 to rows 5‑6, duplicating the associated comments with ShapeCollection.CopyCommentsInRange, and validating that the copied comments retain the original note text and correct row indices.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ---------- Add original comments ----------
            // Comment in cell A1 (row 0, column 0)
            int commentIdx1 = worksheet.Comments.Add(0, 0);
            Comment comment1 = worksheet.Comments[commentIdx1];
            comment1.Note = "Original comment in A1";

            // Comment in cell A2 (row 1, column 0)
            int commentIdx2 = worksheet.Comments.Add(1, 0);
            Comment comment2 = worksheet.Comments[commentIdx2];
            comment2.Note = "Original comment in A2";

            // ---------- Define source range ----------
            CellArea sourceArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 1,
                EndColumn = 0   // only column A
            };

            // Destination start row and column (copy to rows 5 and 6, column A)
            int destStartRow = 5;
            int destStartColumn = 0;

            // ---------- Copy rows (data & formats) ----------
            // Copy rows 0-1 to rows 5-6
            worksheet.Cells.CopyRows(worksheet.Cells, sourceArea.StartRow, destStartRow, sourceArea.EndRow - sourceArea.StartRow + 1);

            // ---------- Copy comments ----------
            // ShapeCollection.CopyCommentsInRange copies comments from source range to destination range
            ShapeCollection shapes = worksheet.Shapes;
            shapes.CopyCommentsInRange(shapes, sourceArea, destStartRow, destStartColumn);

            // ---------- Validation ----------
            // Expected notes after copy
            string[] expectedNotes = { "Original comment in A1", "Original comment in A2" };
            bool allValid = true;

            for (int i = 0; i < expectedNotes.Length; i++)
            {
                int srcRow = sourceArea.StartRow + i;
                int destRow = destStartRow + i;

                // Retrieve comment from destination cell
                Comment destComment = worksheet.Comments[destRow, destStartColumn];

                if (destComment == null)
                {
                    Console.WriteLine($"Comment missing at destination cell ({destRow}, {destStartColumn}).");
                    allValid = false;
                    continue;
                }

                // Compare note text
                if (destComment.Note == expectedNotes[i])
                {
                    Console.WriteLine($"Comment correctly copied to row {destRow}: \"{destComment.Note}\"");
                }
                else
                {
                    Console.WriteLine($"Comment mismatch at row {destRow}. Expected \"{expectedNotes[i]}\", found \"{destComment.Note}\"");
                    allValid = false;
                }

                // Additional check: verify the Row property matches the destination row
                if (destComment.Row != destRow)
                {
                    Console.WriteLine($"Row property mismatch: comment.Row = {destComment.Row}, expected {destRow}");
                    allValid = false;
                }
            }

            Console.WriteLine(allValid ? "All comments transferred correctly." : "Comment transfer validation failed.");

            // Save the workbook (optional, just to visualize the result)
            workbook.Save("CommentCopyValidationResult.xlsx");
        }
    }
}
