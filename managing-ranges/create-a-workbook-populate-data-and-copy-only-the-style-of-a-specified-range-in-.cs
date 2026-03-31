using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsStyleCopyDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create source workbook and populate data --------------------
            Workbook sourceWorkbook = new Workbook();                     // create a new workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];        // get the first worksheet
            sourceSheet.Name = "Source";

            // Fill sample data in the source range A1:D5
            Cells srcCells = sourceSheet.Cells;
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    srcCells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a style and apply it to the source range
            Style srcStyle = sourceWorkbook.CreateStyle();
            srcStyle.Font.Name = "Arial";
            srcStyle.Font.Size = 14;
            srcStyle.Font.IsBold = true;
            srcStyle.ForegroundColor = Color.Yellow;
            srcStyle.Pattern = BackgroundType.Solid;

            // Define the source range and set its style
            AsposeRange srcRange = srcCells.CreateRange("A1:D5");
            srcRange.SetStyle(srcStyle);

            // -------------------- Create destination workbook and copy only style --------------------
            Workbook destWorkbook = new Workbook();                       // create a new workbook for destination
            int destSheetIndex = destWorkbook.Worksheets.Add();
            Worksheet destSheet = destWorkbook.Worksheets[destSheetIndex];
            destSheet.Name = "Destination";

            // Fill different data in the destination range (style will be copied later)
            Cells destCells = destSheet.Cells;
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    destCells[row, col].PutValue($"Dest {row + 1},{col + 1}");
                }
            }

            // Define the destination range (same size as source)
            AsposeRange destRange = destCells.CreateRange("A1:D5");

            // Copy only the style from source range to destination range
            destRange.CopyStyle(srcRange);

            // -------------------- Save the result --------------------
            destWorkbook.Save("DestinationWithCopiedStyle.xlsx");
        }
    }
}