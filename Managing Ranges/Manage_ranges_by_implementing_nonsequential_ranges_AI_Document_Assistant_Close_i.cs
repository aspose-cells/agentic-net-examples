using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data in two separate blocks.
        cells["A1"].PutValue("Item1");
        cells["A2"].PutValue(10);
        cells["B1"].PutValue("Item2");
        cells["B2"].PutValue(20);

        cells["D4"].PutValue("Item3");
        cells["D5"].PutValue(30);
        cells["E4"].PutValue("Item4");
        cells["E5"].PutValue(40);

        // Create the first non‑contiguous range (A1:B2) and give it a name.
        AsposeRange range1 = cells.CreateRange("A1", "B2");
        range1.Name = "FirstBlock";

        // Create the second non‑contiguous range (D4:E5) and give it a name.
        AsposeRange range2 = cells.CreateRange("D4", "E5");
        range2.Name = "SecondBlock";

        // Define a simple style (light yellow background) to apply to both ranges.
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightYellow;
        style.Pattern = BackgroundType.Solid;
        StyleFlag flag = new StyleFlag();
        flag.CellShading = true;

        // Apply the style to each range.
        range1.ApplyStyle(style, flag);
        range2.ApplyStyle(style, flag);

        // Save the workbook in XLSX format.
        workbook.Save("NonSequentialRanges.xlsx");
    }
}