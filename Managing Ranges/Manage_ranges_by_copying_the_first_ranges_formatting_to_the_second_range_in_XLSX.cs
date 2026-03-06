using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class RangeCopyFormattingDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (source)
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Define a source range and apply some formatting
            Cells sourceCells = sourceSheet.Cells;
            AsposeRange sourceRange = sourceCells.CreateRange("A1:D5");
            Style srcStyle = workbook.CreateStyle();
            srcStyle.Font.Name = "Calibri";
            srcStyle.Font.Size = 12;
            srcStyle.Font.IsBold = true;
            srcStyle.ForegroundColor = Color.LightGreen;
            srcStyle.Pattern = BackgroundType.Solid;
            sourceRange.SetStyle(srcStyle);

            // Add a second worksheet (destination)
            int destIndex = workbook.Worksheets.Add();
            Worksheet destSheet = workbook.Worksheets[destIndex];
            destSheet.Name = "Destination";

            // Define a destination range with the same size
            Cells destCells = destSheet.Cells;
            AsposeRange destRange = destCells.CreateRange("A1:D5");

            // Copy formatting from source range to destination range
            destRange.CopyStyle(sourceRange);

            // Save the workbook in XLSX format
            workbook.Save("RangeCopyFormattingDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            RangeCopyFormattingDemo.Run();
        }
    }
}