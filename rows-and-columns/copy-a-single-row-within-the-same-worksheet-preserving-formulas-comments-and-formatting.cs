// Title: Copy a Worksheet Row with Formulas, Comments, and Formatting – Aspose.Cells C# Example
// Description: Shows how to duplicate a single row in the same worksheet using Aspose.Cells for .NET while keeping cell values, formulas, styles, and attached comments. The example calls Cells.CopyRow for data and formatting, then copies comments manually to the destination row before saving the workbook.
// Keywords: Aspose.Cells copy row C# | duplicate worksheet row .NET | preserve formulas Aspose.Cells | copy cell comments Aspose.Cells | row formatting copy Aspose.Cells | Cells.CopyRow example | C# Excel row cloning
// Common Searches: Aspose.Cells copy row preserving formulas | How to copy a row with comments in Aspose.Cells C# | Copy row formatting Aspose.Cells .NET | Duplicate Excel row using Aspose.Cells | CopyRow method example Aspose.Cells
// Developer Intent: Duplicate a row in the same sheet without losing formulas, styles, or comments.
// Use Cases: Replicate a header row that contains calculations and explanatory notes for a new report section. | Generate multiple data‑entry rows from a template row that includes validation, formulas, and comments. | Move a calculation row to another part of the worksheet while preserving its appearance and attached remarks.
// AI Prompts: Write C# code that copies a row with Aspose.Cells and automatically transfers all comments to the new row. | Explain why Cells.CopyRow does not copy comments and propose a utility method to clone them efficiently. | Provide a robust copy‑row routine that handles missing comment authors and preserves merged cells and data validation.

using Aspose.Cells;
using System;

// Shows how to duplicate a single row in the same worksheet using Aspose.Cells for .NET while keeping cell values, formulas, styles, and attached comments. The example calls Cells.CopyRow for data and formatting, then copies comments manually to the destination row before saving the workbook.
class CopyRowExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate source row (row index 0) with values, a formula, and a comment
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);
        cells["C1"].Formula = "=A1+B1";

        // Add a comment to cell A1
        int commentIdx = worksheet.Comments.Add(0, 0); // row 0, column 0
        Comment srcComment = worksheet.Comments[commentIdx];
        srcComment.Note = "Sample comment on source row";

        // Copy the entire row 0 to row 2 (preserves data, formulas, and formatting)
        cells.CopyRow(cells, 0, 2);

        // Manually copy comments from the source row to the destination row
        foreach (Comment comment in worksheet.Comments)
        {
            if (comment.Row == 0) // source row
            {
                int destRow = 2; // destination row index
                int col = comment.Column;

                // Add a new comment at the same column in the destination row
                int newCommentIdx = worksheet.Comments.Add(destRow, col);
                Comment destComment = worksheet.Comments[newCommentIdx];
                destComment.Note = comment.Note;
                destComment.Author = comment.Author;
            }
        }

        // Save the workbook
        workbook.Save("CopyRowResult.xlsx");
    }
}
