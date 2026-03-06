using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class CopyRangeFormattingDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Source worksheet ----------
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Create a source range and apply a style to it
            Cells srcCells = srcSheet.Cells;
            AsposeRange srcRange = srcCells.CreateRange("A1:C3");

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

            // Create a destination range of the same size
            Cells destCells = destSheet.Cells;
            AsposeRange destRange = destCells.CreateRange("A1:C3");

            // Copy formatting from the source range to the destination range
            destRange.CopyStyle(srcRange);

            // Save the workbook to an XLSX file
            workbook.Save("CopyRangeFormattingDemo.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CopyRangeFormattingDemo.Run();
        }
    }
}