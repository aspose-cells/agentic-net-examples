// Title: Validate Comment Transfer When Copying Rows with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds comments to cells A1 and B2, copies rows 0‑1 to rows 5‑6 using Cells.CopyRows, duplicates the comments with ShapeCollection.CopyCommentsInRange, and verifies that the notes and row indices are preserved before saving the file.
// Keywords: Aspose.Cells | C# | Cells.CopyRows | ShapeCollection.CopyCommentsInRange | copy rows with comments | comment validation | Excel comment duplication | row copy integrity
// Common Searches: Aspose.Cells copy rows keep comments | Validate copied comments in .NET Excel | ShapeCollection.CopyCommentsInRange usage | How to verify comment transfer after row copy | C# example for copying rows with comments
// Developer Intent: Confirm that comments attached to source rows are accurately reproduced on the destination rows after a row‑copy operation.
// Use Cases: Add cell comments, copy a block of rows, and ensure comments move with the data. | Automate quality checks for comment integrity in generated reports. | Persist validated comment placement when saving the workbook to disk.
// AI Prompts: Generate C# code that copies rows with Aspose.Cells and asserts that all cell comments are duplicated correctly. | Explain step‑by‑step how ShapeCollection.CopyCommentsInRange transfers comments during a row copy. | Create a unit test in C# that checks comment text and row index after using Cells.CopyRows and CopyCommentsInRange.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds comments to cells A1 and B2, copies rows 0‑1 to rows 5‑6 using Cells.CopyRows, duplicates the comments with ShapeCollection.CopyCommentsInRange, and verifies that the notes and row indices are preserved before saving the file.
    public class CommentsCopyValidation
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Add comments to original rows (rows 0 and 1)
                // -------------------------------------------------
                // Comment in cell A1 (row 0, column 0)
                int commentIdx1 = worksheet.Comments.Add(0, 0);
                Comment comment1 = worksheet.Comments[commentIdx1];
                comment1.Note = "Comment on original row 0";

                // Comment in cell B2 (row 1, column 1)
                int commentIdx2 = worksheet.Comments.Add(1, 1);
                Comment comment2 = worksheet.Comments[commentIdx2];
                comment2.Note = "Comment on original row 1";

                // -------------------------------------------------
                // 2. Copy rows 0-1 to destination starting at row 5
                // -------------------------------------------------
                // Copy the row data and formats
                worksheet.Cells.CopyRows(worksheet.Cells, 0, 5, 2); // copies 2 rows (0 and 1) to rows 5 and 6

                // Copy the comments associated with the source range to the destination range
                ShapeCollection shapes = worksheet.Shapes;
                CellArea sourceArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 1,
                    EndColumn = 1
                };
                // Destination starts at row 5, column 0
                shapes.CopyCommentsInRange(shapes, sourceArea, 5, 0);

                // -------------------------------------------------
                // 3. Validate that comments were transferred correctly
                // -------------------------------------------------
                // Expected destination cells: A6 (row 5, col 0) and B7 (row 6, col 1)
                Comment destComment1 = worksheet.Comments["A6"];
                Comment destComment2 = worksheet.Comments["B7"];

                bool isFirstCommentCorrect = destComment1 != null && destComment1.Note == comment1.Note;
                bool isSecondCommentCorrect = destComment2 != null && destComment2.Note == comment2.Note;

                Console.WriteLine("First comment transferred correctly: " + isFirstCommentCorrect);
                Console.WriteLine("Second comment transferred correctly: " + isSecondCommentCorrect);

                // Additional sanity check: row indices reported by the Comment objects
                if (destComment1 != null)
                    Console.WriteLine("Destination comment 1 row index: " + destComment1.Row); // should be 5
                if (destComment2 != null)
                    Console.WriteLine("Destination comment 2 row index: " + destComment2.Row); // should be 6

                // -------------------------------------------------
                // 4. Save the workbook (lifecycle rule)
                // -------------------------------------------------
                string outputPath = "CommentsCopyValidationResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point required for compilation
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
