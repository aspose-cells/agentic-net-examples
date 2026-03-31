using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace NonSequentialRangeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string? inputFile = null; // e.g., "C:\\input.xlsx";
            string outputFile = Path.Combine(Path.GetTempPath(), "NonSequentialRanges.xlsx");

            ManageNonSequentialRanges(inputFile, outputFile);
            Console.WriteLine($"Workbook saved to: {outputFile}");
        }

        static void ManageNonSequentialRanges(string? inputFile, string outputFile)
        {
            Workbook workbook = string.IsNullOrEmpty(inputFile) ? new Workbook() : new Workbook(inputFile);
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["D4"].PutValue("Total");
            sheet.Cells["E4"].PutValue(30);
            sheet.Cells["D5"].PutValue("Average");
            sheet.Cells["E5"].PutValue(15);

            AsposeRange range1 = sheet.Cells.CreateRange("A1:B3");
            AsposeRange range2 = sheet.Cells.CreateRange("D4:E5");

            // Apply styles
            Style style1 = workbook.CreateStyle();
            style1.ForegroundColor = Color.LightBlue;
            style1.Pattern = BackgroundType.Solid;
            StyleFlag flag1 = new StyleFlag { CellShading = true };
            range1.ApplyStyle(style1, flag1);

            Style style2 = workbook.CreateStyle();
            style2.ForegroundColor = Color.LightGreen;
            style2.Pattern = BackgroundType.Solid;
            StyleFlag flag2 = new StyleFlag { CellShading = true };
            range2.ApplyStyle(style2, flag2);

            workbook.Save(outputFile);
        }
    }
}