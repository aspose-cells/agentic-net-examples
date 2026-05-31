using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ValidateCommentCopyOnRowCopy
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Add original comments ----------
                // Comment in cell A1 (row 0, column 0)
                int idxA1 = worksheet.Comments.Add(0, 0);
                Comment commentA1 = worksheet.Comments[idxA1];
                commentA1.Note = "Comment on original row 1";

                // Comment in cell B2 (row 1, column 1)
                int idxB2 = worksheet.Comments.Add(1, 1);
                Comment commentB2 = worksheet.Comments[idxB2];
                commentB2.Note = "Comment on original row 2";

                // Store original notes for later comparison
                string originalNoteA1 = commentA1.Note;
                string originalNoteB2 = commentB2.Note;

                // ---------- Define source range ----------
                // We will copy rows 0 and 1 (A1:B2 area)
                CellArea sourceArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 1,
                    EndColumn = 1
                };

                // ---------- Destination start position ----------
                // Copy to rows starting at index 5 (Excel row 6), column 0 (A column)
                int destStartRow = 5;
                int destStartColumn = 0;

                // ---------- Copy rows ----------
                // This copies cell values, formats, etc.
                worksheet.Cells.CopyRows(worksheet.Cells, sourceArea.StartRow, destStartRow,
                    sourceArea.EndRow - sourceArea.StartRow + 1);

                // ---------- Copy comments ----------
                // ShapeCollection.CopyCommentsInRange copies comments from the source range to the destination.
                worksheet.Shapes.CopyCommentsInRange(worksheet.Shapes, sourceArea, destStartRow, destStartColumn);

                // ---------- Validate copied comments ----------
                // Destination cells corresponding to original A1 and B2 are A6 and B7 respectively.
                Comment copiedCommentA6 = worksheet.Comments["A6"]; // Row 5, Column 0
                Comment copiedCommentB7 = worksheet.Comments["B7"]; // Row 6, Column 1

                bool a6Matches = copiedCommentA6 != null && copiedCommentA6.Note == originalNoteA1;
                bool b7Matches = copiedCommentB7 != null && copiedCommentB7.Note == originalNoteB2;

                Console.WriteLine("Validation Results:");
                Console.WriteLine($"Comment copied to A6 matches original: {a6Matches}");
                Console.WriteLine($"Comment copied to B7 matches original: {b7Matches}");

                // ---------- Save the workbook ----------
                workbook.Save("CommentCopyValidation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateCommentCopyOnRowCopy.Run();
        }
    }
}