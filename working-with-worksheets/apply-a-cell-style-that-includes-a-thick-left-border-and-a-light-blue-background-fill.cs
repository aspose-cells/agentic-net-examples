using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsStyleExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Choose a cell to style (e.g., B2)
            Cell cell = cells["B2"];
            cell.PutValue("Styled Cell");

            // Create a new style
            Style style = workbook.CreateStyle();

            // Set a solid fill with light blue background
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightBlue;

            // Set a thick left border (color can be chosen as needed)
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thick, Color.Black);

            // Create a style flag to apply only the left border and cell shading
            StyleFlag flag = new StyleFlag
            {
                LeftBorder = true,   // Apply left border settings
                CellShading = true   // Apply background fill settings
            };

            // Apply the style to the cell using the flag
            cell.SetStyle(style, flag);

            // Save the workbook
            workbook.Save("StyledCell.xlsx");
        }
    }
}