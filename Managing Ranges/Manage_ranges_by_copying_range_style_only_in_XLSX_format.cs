using System;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

class CopyRangeStyleDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ---------- Source worksheet ----------
        Worksheet srcSheet = workbook.Worksheets[0];
        srcSheet.Name = "Source";

        // Create a source range and apply a style to it
        Cells srcCells = srcSheet.Cells;
        AsposeRange srcRange = srcCells.CreateRange("A1:D5");
        Style srcStyle = workbook.CreateStyle();
        srcStyle.Font.Name = "Arial";
        srcStyle.Font.Size = 14;
        srcStyle.Font.IsBold = true;
        srcStyle.ForegroundColor = Color.Yellow;
        srcStyle.Pattern = BackgroundType.Solid;
        srcRange.SetStyle(srcStyle);

        // ---------- Destination worksheet ----------
        int destIndex = workbook.Worksheets.Add(); // add a new worksheet
        Worksheet destSheet = workbook.Worksheets[destIndex];
        destSheet.Name = "Destination";

        // Create a destination range and copy only the style from the source range
        Cells destCells = destSheet.Cells;
        AsposeRange destRange = destCells.CreateRange("A1:D5");
        destRange.CopyStyle(srcRange); // copies style settings only

        // Save the workbook in XLSX format
        workbook.Save("RangeCopyStyleDemo.xlsx", SaveFormat.Xlsx);
    }
}