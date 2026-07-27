using System;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a source workbook and add some data with style
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";

            // Populate cells
            sourceSheet.Cells["A1"].PutValue("Styled Text");
            sourceSheet.Cells["A2"].PutValue(123);

            // Apply a style to A1
            Style style = sourceWorkbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 14;
            style.Font.IsBold = true;
            style.ForegroundColor = System.Drawing.Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            sourceSheet.Cells["A1"].SetStyle(style);

            // Create a destination workbook (empty)
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "DestinationSheet";

            // Copy the source worksheet to the destination worksheet, preserving styles
            destSheet.Copy(sourceSheet);

            // Save the destination workbook
            destWorkbook.Save("CopiedWorksheet.xlsx", SaveFormat.Xlsx);
        }
    }
}