using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class TransferRowsWithComments
{
    static void Main()
    {
        // ---------- Create source workbook and add data/comments ----------
        Workbook srcWb = new Workbook();
        Worksheet srcWs = srcWb.Worksheets[0];

        // Add sample data in rows 0‑2 (A1‑A3)
        srcWs.Cells["A1"].PutValue("Row1");
        srcWs.Cells["A2"].PutValue("Row2");
        srcWs.Cells["A3"].PutValue("Row3");

        // Add comments to A1 and A3
        int idx1 = srcWs.Comments.Add("A1");
        srcWs.Comments[idx1].Note = "Comment on Row1";

        int idx2 = srcWs.Comments.Add("A3");
        srcWs.Comments[idx2].Note = "Comment on Row3";

        // ---------- Create destination workbook ----------
        Workbook destWb = new Workbook();
        Worksheet destWs = destWb.Worksheets[0];

        // ---------- Define source range (rows 0‑2, column 0) ----------
        CellArea srcArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 2,
            EndColumn = 0
        };

        // ---------- Copy rows from source to destination ----------
        // Copy 3 rows starting at source row 0 to destination row 5
        destWs.Cells.CopyRows(srcWs.Cells, srcArea.StartRow, 5, srcArea.EndRow - srcArea.StartRow + 1);

        // ---------- Copy comments from source range to destination ----------
        // Destination start row = 5, start column = 0 (A column)
        srcWs.Shapes.CopyCommentsInRange(srcWs.Shapes, srcArea, 5, 0);

        // ---------- Verify that comments were copied ----------
        // Row index 5 corresponds to cell A6, row index 7 corresponds to A8
        Comment destComment1 = destWs.Comments["A6"];
        Comment destComment2 = destWs.Comments["A8"];

        Console.WriteLine("Comment at A6: " + (destComment1 != null ? destComment1.Note : "None"));
        Console.WriteLine("Comment at A8: " + (destComment2 != null ? destComment2.Note : "None"));

        // ---------- Save workbooks ----------
        srcWb.Save("Source.xlsx");
        destWb.Save("Destination.xlsx");
    }
}