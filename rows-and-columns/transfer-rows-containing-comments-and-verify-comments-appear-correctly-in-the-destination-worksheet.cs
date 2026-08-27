// Title: Copy rows with cell comments from one worksheet to another and verify comment placement using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy a range of rows that contain cell comments from a source worksheet to a specific start row in a destination worksheet, preserving the original comment locations. | After copying, retrieve the comments from the destination worksheet and programmatically confirm that each comment's text matches the source and appears in the expected cells.
// Common Searches: asp.net copy rows with comments using Aspose.Cells | preserve Excel cell comments when copying rows in C# Aspose.Cells | verify comment locations after row transfer Aspose.Cells .NET | copy rows 2-3 to row 6 with comments Aspose.Cells example | Aspose.Cells copy rows between worksheets while keeping comments
// Tags: copy rows with comments Aspose.Cells | preserve cell comments Aspose.Cells | verify comment placement C# | transfer rows between worksheets Aspose.Cells | Aspose.Cells comment copy verification

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace TransferRowsWithComments
{
    // The example creates a source workbook, adds data and cell comments, then copies rows 2‑3 (including their comments) to row 6 of a destination worksheet using Aspose.Cells for .NET. It retrieves the comments at A6 and B7 in the destination sheet to confirm they were transferred correctly, and saves both workbooks.
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook and add data/comments ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Add sample data
            sourceSheet.Cells["A2"].PutValue("Row 2 Data");
            sourceSheet.Cells["B2"].PutValue(200);
            sourceSheet.Cells["A3"].PutValue("Row 3 Data");
            sourceSheet.Cells["B3"].PutValue(300);

            // Add comments to the rows we will copy
            int commentIdx1 = sourceSheet.Comments.Add("A2");
            sourceSheet.Comments[commentIdx1].Note = "Comment on A2";

            int commentIdx2 = sourceSheet.Comments.Add("B3");
            sourceSheet.Comments[commentIdx2].Note = "Comment on B3";

            // Define the source range that includes the rows with comments (rows 2-3)
            CellArea sourceArea = new CellArea
            {
                StartRow = 1,      // Row index 1 => A2
                StartColumn = 0,   // Column A
                EndRow = 2,        // Row index 2 => A3
                EndColumn = 1      // Column B
            };

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "Destination";

            // Copy the rows (including data and formats) from source to destination
            // Destination start row = 5 (i.e., row 6 in Excel)
            destSheet.Cells.CopyRows(sourceSheet.Cells, sourceArea.StartRow, 5, sourceArea.EndRow - sourceArea.StartRow + 1);

            // Copy the comments that belong to the source range to the destination range
            // Destination start column = 0 (column A)
            sourceSheet.Shapes.CopyCommentsInRange(sourceSheet.Shapes, sourceArea, 5, 0);

            // ---------- Verify that comments were copied correctly ----------
            // Expected locations after copy: A6 (row index 5) and B7 (row index 6)
            Comment copiedComment1 = destSheet.Comments["A6"];
            Comment copiedComment2 = destSheet.Comments["B7"];

            Console.WriteLine("Copied comment at A6: " + (copiedComment1 != null ? copiedComment1.Note : "Not found"));
            Console.WriteLine("Copied comment at B7: " + (copiedComment2 != null ? copiedComment2.Note : "Not found"));

            // ---------- Save workbooks ----------
            sourceWorkbook.Save("SourceWorkbook.xlsx");
            destWorkbook.Save("DestinationWorkbook.xlsx");
        }
    }
}
