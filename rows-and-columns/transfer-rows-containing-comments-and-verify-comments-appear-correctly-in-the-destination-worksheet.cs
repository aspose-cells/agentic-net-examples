// Title: Copy rows and preserve cell comments with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a range of rows from one worksheet to another using Aspose.Cells, transfer the associated cell comments with Shapes.CopyCommentsInRange, and verify that the comments are retained in the destination sheet.
// Keywords: Aspose.Cells copy rows | copy rows with comments C# | transfer worksheet rows Aspose.Cells | preserve cell comments Aspose | Shapes.CopyCommentsInRange | Aspose.Cells .NET example | Excel row copy with comments
// Common Searches: Aspose.Cells copy rows and keep comments | C# copy Excel rows preserving comments | How to transfer cell comments when copying rows in Aspose.Cells | Copy rows between worksheets Aspose.Cells .NET | Copy rows with comments using Aspose.Cells API
// Developer Intent: Copy selected rows from a source worksheet to a destination worksheet while automatically moving any cell comments that belong to those rows.
// Use Cases: Migrate annotated data rows from a template workbook to a reporting workbook. | Duplicate a block of rows with reviewer notes for further analysis in a separate sheet. | Create a consolidated summary by copying rows that contain comments from multiple source sheets.
// AI Prompts: Generate C# code that copies rows 5‑10 from Sheet1 to Sheet2 at row 20 and also copies all comments within those rows using Aspose.Cells. | Explain the purpose of Shapes.CopyCommentsInRange and show how to adapt it for a multi‑column range. | Write a method that validates that comments were successfully transferred after copying rows with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to copy a range of rows from one worksheet to another using Aspose.Cells, transfer the associated cell comments with Shapes.CopyCommentsInRange, and verify that the comments are retained in the destination sheet.
class TransferRowsWithComments
{
    static void Main()
    {
        // ---------- Create source workbook and add data ----------
        Workbook srcWb = new Workbook();
        Worksheet srcWs = srcWb.Worksheets[0];

        // Populate three rows with sample data
        srcWs.Cells["A1"].PutValue("Row 1 data");
        srcWs.Cells["A2"].PutValue("Row 2 data");
        srcWs.Cells["A3"].PutValue("Row 3 data");

        // Add comments to rows that contain data (A1 and A3)
        int commentIdx1 = srcWs.Comments.Add("A1");
        srcWs.Comments[commentIdx1].Note = "Comment on Row 1";

        int commentIdx2 = srcWs.Comments.Add("A3");
        srcWs.Comments[commentIdx2].Note = "Comment on Row 3";

        // ---------- Create destination workbook ----------
        Workbook destWb = new Workbook();
        Worksheet destWs = destWb.Worksheets[0];

        // ---------- Define source range that includes the rows to copy ----------
        // Here we copy rows 0‑2 (A1:A3). EndColumn is set to 0 because we only need column A.
        CellArea sourceArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 2,
            EndColumn = 0
        };

        // ---------- Copy rows ----------
        // Copy the three rows from source worksheet to destination worksheet,
        // starting at destination row index 5 (i.e., row 6 in Excel terms).
        destWs.Cells.CopyRows(srcWs.Cells, sourceRowIndex: 0, destinationRowIndex: 5, rowNumber: 3);

        // ---------- Copy comments ----------
        // ShapeCollection.CopyCommentsInRange copies all comments that lie inside the
        // specified source range to the destination range (destRow, destColumn).
        srcWs.Shapes.CopyCommentsInRange(srcWs.Shapes, sourceArea, destRow: 5, destColumn: 0);

        // ---------- Verify that comments were copied ----------
        Console.WriteLine("Comments present in destination worksheet:");
        foreach (Comment comment in destWs.Comments)
        {
            // Comment.Row gives the zero‑based row index where the comment resides.
            // Comment.Column gives the zero‑based column index.
            Console.WriteLine($"- Cell {CellsHelper.CellIndexToName(comment.Row, comment.Column)}: {comment.Note}");
        }

        // ---------- Save workbooks (optional) ----------
        srcWb.Save("SourceWorkbook.xlsx");
        destWb.Save("DestinationWorkbook.xlsx");
    }
}
