using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            ApplyRangeFormatting.Run();
        }
    }

    public class ApplyRangeFormatting
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Source worksheet ----------
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Create a source range and apply a custom style
            Cells srcCells = srcSheet.Cells;
            var srcRange = srcCells.CreateRange("A1:D5");
            Style srcStyle = workbook.CreateStyle();
            srcStyle.Font.Name = "Calibri";
            srcStyle.Font.Size = 12;
            srcStyle.Font.IsBold = true;
            srcStyle.ForegroundColor = Color.LightGreen;
            srcStyle.Pattern = BackgroundType.Solid;
            srcRange.SetStyle(srcStyle);

            // ---------- Destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            destSheet.Name = "Destination";

            // Create a destination range (same size as source)
            Cells destCells = destSheet.Cells;
            var destRange = destCells.CreateRange("A1:D5");

            // Copy formatting from source range to destination range
            destRange.CopyStyle(srcRange);

            // Save the workbook
            workbook.Save("FormattedRangeCopy.xlsx");
        }
    }
}