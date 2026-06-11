using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell
            Cell cell = sheet.Cells["B2"];
            cell.PutValue("Italic with thin top border");

            // Create a style and configure italic font and top border
            Style style = workbook.CreateStyle();
            style.Font.IsItalic = true;                                   // italic font
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; // thin top border
            style.Borders[BorderType.TopBorder].Color = Color.Black;      // optional border color

            // Define which style attributes to apply
            StyleFlag flag = new StyleFlag();
            flag.FontItalic = true;   // apply italic setting
            flag.TopBorder = true;    // apply top border setting

            // Apply the style to the cell using the flag
            cell.SetStyle(style, flag);

            // Save the workbook
            workbook.Save("CustomStyle.xlsx");
        }
    }
}