// Title: Check comment transfer after row copy with Aspose.Cells for .NET
// Description: Shows how to add comments to cells A1 and B2, duplicate rows 0‑1 to rows 5‑6 using Cells.CopyRows, copy the associated comments with ShapeCollection.CopyCommentsInRange, and verify that cells A6 and B7 contain identical notes before saving the workbook.
// Keywords: Aspose.Cells copy rows comments | ShapeCollection.CopyCommentsInRange C# | validate copied comments .NET | row duplication comment preservation | Aspose.Cells comment verification
// Common Searches: Aspose.Cells verify comments after copying rows | Copy comments with ShapeCollection.CopyCommentsInRange example | C# check comment note after row duplication | How to preserve cell comments when using Cells.CopyRows | Validate comment transfer in Aspose.Cells workbook
// Developer Intent: Confirm that comments attached to the original rows appear unchanged in the duplicated rows.
// Use Cases: After using Cells.CopyRows, call ShapeCollection.CopyCommentsInRange to replicate comments to the new row range. | Programmatically inspect the destination cells (A6, B7) to ensure the Comment objects exist and their Note properties match the source comments. | Persist the workbook after validation for downstream processing or manual review.
// AI Prompts: Write C# code that copies rows and their comments from one area to another with Aspose.Cells and asserts that the comment texts are identical. | Create an MSTest unit test that verifies the Note values of comments copied by ShapeCollection.CopyCommentsInRange match the originals. | Explain the required parameters and typical workflow of ShapeCollection.CopyCommentsInRange for preserving comments during row duplication.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentCopyValidation
{
    // Shows how to add comments to cells A1 and B2, duplicate rows 0‑1 to rows 5‑6 using Cells.CopyRows, copy the associated comments with ShapeCollection.CopyCommentsInRange, and verify that cells A6 and B7 contain identical notes before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Add comments to the original rows (A1 and B2)
            // -------------------------------------------------
            CommentCollection comments = worksheet.Comments;

            // Add comment to cell A1 (row 0, column 0)
            int commentIdx1 = comments.Add(0, 0);
            Comment comment1 = comments[commentIdx1];
            comment1.Note = "Original comment in A1";

            // Add comment to cell B2 (row 1, column 1)
            int commentIdx2 = comments.Add(1, 1);
            Comment comment2 = comments[commentIdx2];
            comment2.Note = "Original comment in B2";

            // -------------------------------------------------
            // 2. Copy rows 0-1 to rows 5-6 (preserve data)
            // -------------------------------------------------
            // Copy two rows starting from source row 0 to destination row 5
            worksheet.Cells.CopyRows(worksheet.Cells, 0, 5, 2);

            // -------------------------------------------------
            // 3. Copy comments from source range to destination range
            // -------------------------------------------------
            // Define the source range (rows 0-1, columns 0-1)
            CellArea sourceArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 1,
                EndColumn = 1
            };

            // Get the shapes collection (required by the API)
            ShapeCollection shapes = worksheet.Shapes;

            // Copy comments from the source range to the destination starting at row 5, column 0
            shapes.CopyCommentsInRange(shapes, sourceArea, 5, 0);

            // -------------------------------------------------
            // 4. Validate that comments were transferred correctly
            // -------------------------------------------------
            // Destination cells: A6 (row 5, col 0) and B7 (row 6, col 1)
            Comment destComment1 = worksheet.Comments["A6"];
            Comment destComment2 = worksheet.Comments["B7"];

            Console.WriteLine("Validation Results:");
            Console.WriteLine($"Comment at A6 exists: {destComment1 != null}");
            Console.WriteLine($"Comment at A6 note matches: {destComment1?.Note == comment1.Note}");
            Console.WriteLine($"Comment at B7 exists: {destComment2 != null}");
            Console.WriteLine($"Comment at B7 note matches: {destComment2?.Note == comment2.Note}");

            // -------------------------------------------------
            // 5. Save the workbook (demonstration purpose)
            // -------------------------------------------------
            workbook.Save("CommentCopyValidation.xlsx");
        }
    }
}
