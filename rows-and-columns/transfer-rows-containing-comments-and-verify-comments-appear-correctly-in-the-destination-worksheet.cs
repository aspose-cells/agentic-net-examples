// Title: Transfer rows with cell comments between worksheets using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a block of rows from a source worksheet to a destination worksheet, preserving values, formatting, and cell comments. The example uses Workbook, Cells.CopyRows, and Shapes.CopyCommentsInRange, then verifies the transferred comments and saves the result as an Excel file.
// Keywords: Aspose.Cells | CopyRows | CopyCommentsInRange | C# | .NET | transfer rows | cell comments | worksheet copy | Excel automation
// Common Searches: Aspose.Cells copy rows with comments C# | how to copy rows and comments between worksheets .NET | CopyRows preserve cell comments Aspose | transfer Excel rows programmatically with comments | copy range of rows and comments using Aspose.Cells
// Developer Intent: Copy selected rows and their associated comments from one worksheet to another while keeping formatting intact.
// Use Cases: Migrate annotated data from a template sheet to a report sheet in a new workbook. | Create a historical snapshot by duplicating rows with comments for version control. | Build a consolidated summary by pulling rows with notes from multiple source sheets.
// AI Prompts: Generate C# code that copies rows 2‑6 from Sheet1 to row 15 in Sheet2 and also copies all comments in that range using Aspose.Cells. | Show how to verify that comments were transferred correctly after using Cells.CopyRows and Shapes.CopyCommentsInRange. | Explain best practices for handling comment objects when moving rows across workbooks with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace TransferRowsWithComments
{
    // Demonstrates how to copy a block of rows from a source worksheet to a destination worksheet, preserving values, formatting, and cell comments. The example uses Workbook, Cells.CopyRows, and Shapes.CopyCommentsInRange, then verifies the transferred comments and saves the result as an Excel file.
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook and add data with comments ----------
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Populate three rows with sample data
            srcSheet.Cells["A1"].PutValue("Row1-Data");
            srcSheet.Cells["B1"].PutValue(10);
            srcSheet.Cells["A2"].PutValue("Row2-Data");
            srcSheet.Cells["B2"].PutValue(20);
            srcSheet.Cells["A3"].PutValue("Row3-Data");
            srcSheet.Cells["B3"].PutValue(30);

            // Add comments to each row (A1, A2, A3)
            int c1 = srcSheet.Comments.Add("A1");
            srcSheet.Comments[c1].Note = "Comment for A1";

            int c2 = srcSheet.Comments.Add("A2");
            srcSheet.Comments[c2].Note = "Comment for A2";

            int c3 = srcSheet.Comments.Add("A3");
            srcSheet.Comments[c3].Note = "Comment for A3";

            // ---------- Define source range that contains the rows to copy ----------
            CellArea srcArea = new CellArea
            {
                StartRow = 0,      // Row 0 (A1)
                StartColumn = 0,   // Column 0 (A)
                EndRow = 2,        // Row 2 (A3)
                EndColumn = 1      // Column 1 (B)
            };

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // ---------- Copy rows (data and formats) ----------
            // Copy rows 0‑2 from source to destination starting at row 5 (index 5)
            int destStartRow = 5;
            int rowsToCopy = srcArea.EndRow - srcArea.StartRow + 1;
            destSheet.Cells.CopyRows(srcSheet.Cells, srcArea.StartRow, destStartRow, rowsToCopy);

            // ---------- Copy comments within the same range ----------
            // Destination column start is same as source (0)
            destSheet.Shapes.CopyCommentsInRange(srcSheet.Shapes, srcArea, destStartRow, srcArea.StartColumn);

            // ---------- Verify that comments were copied ----------
            Console.WriteLine("Comments in destination worksheet:");
            foreach (Comment comment in destSheet.Comments)
            {
                // The comment's Row and Column indicate its position
                string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                Console.WriteLine($"{cellName}: {comment.Note}");
            }

            // ---------- Save the destination workbook ----------
            destWorkbook.Save("DestinationWithComments.xlsx");
        }
    }
}
