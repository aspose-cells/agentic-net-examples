using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsFooterStyle
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a style with light blue fill and a thin bottom border
            Style style = workbook.CreateStyle();
            style.BackgroundColor = Color.LightBlue;          // light blue fill
            style.Pattern = BackgroundType.Solid;             // solid fill pattern
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; // thin bottom border
            style.Borders[BorderType.BottomBorder].Color = Color.Black;            // border color

            // Define which style attributes should be applied
            StyleFlag flag = new StyleFlag();
            flag.BottomBorder = true;   // apply bottom border
            flag.CellShading = true;    // apply fill (background color)

            // Determine the footer row index (for demo, use the row after the last used row)
            int footerRowIndex = cells.MaxDataRow + 1;

            // Apply the style to the entire footer row
            cells.ApplyRowStyle(footerRowIndex, style, flag);

            // Save the workbook
            workbook.Save("FooterStyle.xlsx");
        }
    }
}