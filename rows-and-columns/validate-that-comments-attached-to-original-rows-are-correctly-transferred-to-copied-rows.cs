// Title: Validate Comment Transfer When Copying Rows with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds comments to A1 and B2, copies rows 0‑1 to row 5, uses ShapeCollection.CopyCommentsInRange to move the comments, then checks they appear at A6 and B7 before saving.
// Keywords: Aspose.Cells | CopyRows | CopyCommentsInRange | C# | .NET | comment preservation | row duplication | worksheet comments | CellArea | ShapeCollection
// Common Searches: Aspose.Cells copy rows with comments | CopyCommentsInRange C# example | Validate copied comments Aspose.Cells | How to preserve comments when copying rows | Copy rows and comments Aspose.Cells .NET
// Developer Intent: Confirm that cell comments are retained and correctly positioned after rows are copied.
// Use Cases: Duplicate a range of rows while keeping associated comments for reporting templates. | Migrate data blocks between worksheets and ensure comment integrity. | Automated testing of comment preservation after bulk row operations.
// AI Prompts: Write C# code that iterates through worksheet.Comments to compare original and copied comment texts after using CopyRows and CopyCommentsInRange. | Create an MSTest unit test that asserts the presence and exact content of comments at the destination cells after copying rows with Aspose.Cells. | Explain why CopyRows alone does not copy comments and how ShapeCollection.CopyCommentsInRange resolves this issue.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentCopyValidation
{
    // Creates a workbook, adds comments to A1 and B2, copies rows 0‑1 to row 5, uses ShapeCollection.CopyCommentsInRange to move the comments, then checks they appear at A6 and B7 before saving.
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
            comment1.Note = "Comment on original row 0";

            // Comment in cell B2 (row 1, column 1)
            int commentIdx2 = worksheet.Comments.Add(1, 1);
            Comment comment2 = worksheet.Comments[commentIdx2];
            comment2.Note = "Comment on original row 1";

            // ---------- Define source range ----------
            // Source range covering rows 0-1 and columns 0-1
            CellArea sourceArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 1,
                EndColumn = 1
            };

            // ---------- Copy rows ----------
            // Destination start row (e.g., row 5) and column (0)
            int destStartRow = 5;
            int destStartColumn = 0;

            // Copy the rows' data and formats
            worksheet.Cells.CopyRows(worksheet.Cells, sourceArea.StartRow, destStartRow, sourceArea.EndRow - sourceArea.StartRow + 1);

            // ---------- Copy comments ----------
            // Use ShapeCollection.CopyCommentsInRange to transfer comments
            ShapeCollection shapes = worksheet.Shapes;
            shapes.CopyCommentsInRange(shapes, sourceArea, destStartRow, destStartColumn);

            // ---------- Validation ----------
            // Expected destination cells for the copied comments
            string destCell1 = CellsHelper.CellIndexToName(destStartRow, sourceArea.StartColumn); // A6
            string destCell2 = CellsHelper.CellIndexToName(destStartRow + 1, sourceArea.StartColumn + 1); // B7

            // Retrieve copied comments
            Comment copiedComment1 = worksheet.Comments[destCell1];
            Comment copiedComment2 = worksheet.Comments[destCell2];

            // Verify and output results
            Console.WriteLine($"Original comment at A1: {comment1.Note}");
            Console.WriteLine($"Copied comment at {destCell1}: {(copiedComment1 != null ? copiedComment1.Note : "Not found")}");

            Console.WriteLine($"Original comment at B2: {comment2.Note}");
            Console.WriteLine($"Copied comment at {destCell2}: {(copiedComment2 != null ? copiedComment2.Note : "Not found")}");

            // Save the workbook (lifecycle rule)
            workbook.Save("CommentCopyValidation.xlsx");
        }
    }
}
