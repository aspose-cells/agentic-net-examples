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
            // ---------- Create source workbook and populate data ----------
            Workbook srcWorkbook = new Workbook();                         // create source workbook
            Worksheet srcSheet = srcWorkbook.Worksheets[0];               // get first worksheet

            // Fill sample data in range A1:D5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    srcSheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a style and apply it to the source range
            Style srcStyle = srcWorkbook.CreateStyle();                   // create a new style
            srcStyle.Font.Name = "Arial";
            srcStyle.Font.Size = 12;
            srcStyle.Font.IsBold = true;
            srcStyle.ForegroundColor = Color.Yellow;
            srcStyle.Pattern = BackgroundType.Solid;

            // Define the source range (A1:D5) and set the style
            AsposeRange srcRange = srcSheet.Cells.CreateRange("A1:D5");
            srcRange.SetStyle(srcStyle);                                 // apply style to source range

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();                        // create destination workbook
            Worksheet destSheet = destWorkbook.Worksheets[0];             // get first worksheet

            // Create a destination range of the same size
            AsposeRange destRange = destSheet.Cells.CreateRange("A1:D5");

            // Copy only the style from source range to destination range
            destRange.CopyStyle(srcRange);                                // copy style settings

            // Save the destination workbook (contains data without source values, only style)
            destWorkbook.Save("DestinationWithCopiedStyle.xlsx");
        }
    }
}