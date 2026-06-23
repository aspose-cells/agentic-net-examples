using System;
using Aspose.Cells;
using System.Drawing;

class HeaderStyleAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a style for the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.White;
        headerStyle.ForegroundColor = Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;

        // Define which style attributes to apply (all in this case)
        StyleFlag flag = new StyleFlag { All = true };

        // Apply the style to the first row (index 0)
        cells.ApplyRowStyle(0, headerStyle, flag);

        // Freeze the first row so it stays visible while scrolling
        // Freeze at row index 1 (second row), column 0, freezing 1 row and 0 columns
        worksheet.FreezePanes(1, 0, 1, 0);

        // Save the workbook
        workbook.Save("HeaderStyledAndFrozen.xlsx");
    }
}