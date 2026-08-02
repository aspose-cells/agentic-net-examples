// Title: Copy rows with cell comments using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a block of rows from a source worksheet to a destination worksheet while preserving cell comments. The example creates a source workbook, adds data and comments to A1 and A3, uses Cells.CopyRows to transfer the rows, ShapeCollection.CopyCommentsInRange to move the comments, verifies the comments in the target sheet, and saves both workbooks.
// Keywords: Aspose.Cells copy rows | cell comments transfer .NET | ShapeCollection.CopyCommentsInRange | C# copy worksheet rows with comments | preserve comments when copying rows | Aspose.Cells example C# | verify copied comments
// Common Searches: how to copy rows and keep comments Aspose.Cells | copy cell comments between worksheets C# | Aspose.Cells transfer rows with annotations | verify comments after copying rows .NET | copy rows with comments Aspose.Cells example
// Developer Intent: Move selected rows from one worksheet to another and retain any associated cell comments.
// Use Cases: Clone a template section with its notes into a report sheet. | Aggregate commented rows from multiple source files into a summary workbook. | Migrate data blocks with annotations during a workbook conversion project.
// AI Prompts: Generate C# code that copies a range of rows and all related comments from one worksheet to another using Aspose.Cells, then confirms the comments were transferred. | Explain why ShapeCollection.CopyCommentsInRange must be called after Cells.CopyRows when copying rows that contain comments. | Provide a strategy for copying multi‑column rows with comments, handling possible comment offset adjustments.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to copy a block of rows from a source worksheet to a destination worksheet while preserving cell comments. The example creates a source workbook, adds data and comments to A1 and A3, uses Cells.CopyRows to transfer the rows, ShapeCollection.CopyCommentsInRange to move the comments, verifies the comments in the target sheet, and saves both workbooks.
class TransferRowsWithComments
{
    static void Main()
    {
        // ---------- Create source workbook and add data/comments ----------
        Workbook srcWb = new Workbook();
        Worksheet srcWs = srcWb.Worksheets[0];
        srcWs.Name = "Source";

        // Add sample data in rows 0‑2 (A1‑A3)
        srcWs.Cells["A1"].PutValue("Row1");
        srcWs.Cells["A2"].PutValue("Row2");
        srcWs.Cells["A3"].PutValue("Row3");

        // Add comments to A1 and A3
        int idx1 = srcWs.Comments.Add("A1");
        srcWs.Comments[idx1].Note = "Comment on Row1";

        int idx3 = srcWs.Comments.Add("A3");
        srcWs.Comments[idx3].Note = "Comment on Row3";

        // ---------- Create destination workbook ----------
        Workbook destWb = new Workbook();
        Worksheet destWs = destWb.Worksheets[0];
        destWs.Name = "Destination";

        // ---------- Define source range that contains the rows to copy ----------
        CellArea srcArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 2,
            EndColumn = 0   // only column A is needed for this example
        };

        // ---------- Copy rows data from source to destination ----------
        // CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowNumber)
        destWs.Cells.CopyRows(srcWs.Cells, srcArea.StartRow, 0,
            srcArea.EndRow - srcArea.StartRow + 1);

        // ---------- Copy comments from source range to destination ----------
        // ShapeCollection.CopyCommentsInRange(shapes, ca, destRow, destColumn)
        srcWs.Shapes.CopyCommentsInRange(srcWs.Shapes, srcArea, 0, 0);

        // ---------- Verify that comments were copied ----------
        Comment destCommentA1 = destWs.Comments["A1"];
        Comment destCommentA3 = destWs.Comments["A3"];

        Console.WriteLine("Destination A1 comment: " +
            (destCommentA1 != null ? destCommentA1.Note : "None"));
        Console.WriteLine("Destination A3 comment: " +
            (destCommentA3 != null ? destCommentA3.Note : "None"));

        // ---------- Save workbooks ----------
        srcWb.Save("Source.xlsx");
        destWb.Save("Destination.xlsx");
    }
}
