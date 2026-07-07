using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsColumnStyleExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a style with italic font and light gray fill
            Style columnStyle = workbook.CreateStyle();
            columnStyle.Font.IsItalic = true;                     // italic text
            columnStyle.ForegroundColor = Color.LightGray;        // fill color
            columnStyle.Pattern = BackgroundType.Solid;           // apply fill

            // Specify which style attributes to apply
            StyleFlag flag = new StyleFlag
            {
                FontItalic = true,    // apply italic setting
                CellShading = true    // apply background fill
            };

            // Apply the style to the entire column Q (index 16, zero‑based)
            cells.Columns[16].ApplyStyle(columnStyle, flag);

            // Save the workbook
            workbook.Save("ColumnQ_Styled.xlsx");
        }
    }
}