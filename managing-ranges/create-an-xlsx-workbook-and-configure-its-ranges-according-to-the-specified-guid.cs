using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is XLSX)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data in the worksheet
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["B1"].PutValue(5);
        sheet.Cells["B2"].PutValue(15);
        sheet.Cells["B3"].PutValue(25);

        // Create a range that covers cells A1:B3
        Aspose.Cells.Range dataRange = sheet.Cells.CreateRange("A1", "B3");
        // Assign a name to the range for later reference
        dataRange.Name = "MyData";

        // Apply a light yellow background to the entire range
        Style rangeStyle = workbook.CreateStyle();
        rangeStyle.ForegroundColor = Color.LightYellow;
        rangeStyle.Pattern = BackgroundType.Solid;
        StyleFlag styleFlag = new StyleFlag { CellShading = true };
        dataRange.ApplyStyle(rangeStyle, styleFlag);

        // Use the named range in a formula (sum of all cells in the range)
        sheet.Cells["C1"].Formula = "=SUM(MyData)";

        // Calculate the formula so that the result is stored in C1
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("ConfiguredRanges.xlsx", SaveFormat.Xlsx);
    }
}