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

            // Define the index of the footer row (for example, row 10)
            int footerRowIndex = 9; // zero‑based index (row 10 in Excel)

            // Create a new style
            Style footerStyle = workbook.CreateStyle();

            // Set a light blue fill
            footerStyle.BackgroundColor = Color.LightBlue;
            footerStyle.Pattern = BackgroundType.Solid;

            // Set a thin bottom border
            footerStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            footerStyle.Borders[BorderType.BottomBorder].Color = Color.Black;

            // Create a style flag to indicate which parts of the style should be applied
            StyleFlag flag = new StyleFlag
            {
                Borders = true,          // apply border settings
                CellShading = true       // apply fill (background) settings
            };

            // Apply the style to the entire footer row
            sheet.Cells.ApplyRowStyle(footerRowIndex, footerStyle, flag);

            // Save the workbook
            workbook.Save("FooterStyle.xlsx");
        }
    }
}