using System;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is XLSX)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate source range A1:C3 with sample data
        AsposeRange srcRange = cells.CreateRange("A1", "C3");
        int counter = 1;
        for (int i = 0; i < srcRange.RowCount; i++)
        {
            for (int j = 0; j < srcRange.ColumnCount; j++)
            {
                srcRange[i, j].PutValue(counter++);
            }
        }

        // Define a style (bold font, yellow background)
        Style style = workbook.CreateStyle();
        style.Font.IsBold = true;
        style.ForegroundColor = Color.Yellow;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to the source range
        srcRange.SetStyle(style);

        // Create destination range E1:G3
        AsposeRange destRange = cells.CreateRange("E1", "G3");

        // Copy only the style from source range to destination range
        destRange.CopyStyle(srcRange);

        // Save the workbook as an XLSX file
        workbook.Save("RangeStyleCopyDemo.xlsx");
    }
}