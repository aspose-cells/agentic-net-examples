using System;
using Aspose.Cells;

namespace AsposeCellsDefaultFontDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ----- Set the workbook's default font -----
            // Get the default style object, modify its font, and assign it back
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.Font.Name = "Arial";
            defaultStyle.Font.Size = 10;
            workbook.DefaultStyle = defaultStyle;

            // ----- Apply the default style to all existing cells -----
            // Create a style that matches the default font settings
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 10;

            // Define which style attributes should be applied
            StyleFlag flag = new StyleFlag();
            flag.FontName = true;
            flag.FontSize = true;

            // Apply the style to the entire first worksheet (all cells)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.ApplyStyle(style, flag);

            // Example: put some values to verify the font is applied
            sheet.Cells["A1"].PutValue("Hello, World!");
            sheet.Cells["B2"].PutValue(12345);

            // Save the workbook to a file
            workbook.Save("DefaultFontDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}