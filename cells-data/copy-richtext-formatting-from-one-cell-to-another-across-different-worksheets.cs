using System;
using System.Drawing;
using Aspose.Cells;

class CopyRichTextDemo
{
    static void Main()
    {
        // Create source workbook and get its first worksheet
        Workbook srcWorkbook = new Workbook();
        Worksheet srcSheet = srcWorkbook.Worksheets[0];

        // Put rich‑text content into source cell A1
        Cell srcCell = srcSheet.Cells["A1"];
        srcCell.PutValue("Rich Text Example");

        // Apply a style to the source cell (font, size, color, bold)
        Style richStyle = srcWorkbook.CreateStyle();
        richStyle.Font.Name = "Arial";
        richStyle.Font.Size = 12;
        richStyle.Font.IsBold = true;
        richStyle.Font.Color = Color.Blue;
        srcCell.SetStyle(richStyle);

        // Create destination workbook and get its first worksheet
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];

        // Destination cell where the content will be copied
        Cell destCell = destSheet.Cells["B2"];

        // Copy the source cell (value and formatting) to the destination cell
        destCell.Copy(srcCell);

        // Save both workbooks
        srcWorkbook.Save("SourceWorkbook.xlsx");
        destWorkbook.Save("DestinationWorkbook.xlsx");
    }
}