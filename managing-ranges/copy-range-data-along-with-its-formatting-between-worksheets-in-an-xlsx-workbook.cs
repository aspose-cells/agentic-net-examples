using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Source worksheet
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Fill source range with data
            Cells srcCells = srcSheet.Cells;
            srcCells["A1"].PutValue("Header");
            srcCells["A2"].PutValue(123);
            srcCells["B2"].PutValue(456);

            // Apply a style to the source range
            Style style = workbook.CreateStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            srcCells.CreateRange("A1:B2").SetStyle(style);

            // Destination worksheet
            Worksheet destSheet = workbook.Worksheets.Add("Destination");
            Cells destCells = destSheet.Cells;

            // Define source and destination ranges
            AsposeRange sourceRange = srcCells.CreateRange("A1:B2");
            AsposeRange destRange = destCells.CreateRange("C5:D6");

            // Set paste options to copy everything (data + formatting)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All
            };

            // Copy source range to destination range preserving formatting
            sourceRange.Copy(destRange, pasteOptions);

            // Save the workbook
            workbook.Save("RangeCopyWithFormatting.xlsx");
        }
    }
}