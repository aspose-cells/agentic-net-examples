using System;
using System.Drawing;
using Aspose.Cells;

class ApplyStyleToColumnQ
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define a style with italic font and light gray fill
        Style style = workbook.CreateStyle();
        style.Font.IsItalic = true;                 // italic text
        style.ForegroundColor = Color.LightGray;    // light gray background
        style.Pattern = BackgroundType.Solid;       // apply background color

        // Specify which style attributes to apply
        StyleFlag flag = new StyleFlag();
        flag.FontItalic = true;    // apply italic setting
        flag.CellShading = true;   // apply background fill

        // Apply the style to the entire column Q (zero‑based index 16)
        cells.Columns[16].ApplyStyle(style, flag);

        // Example data to visualize the applied style
        cells["Q1"].PutValue("Header");
        cells["Q2"].PutValue("Sample Data");

        // Save the workbook
        workbook.Save("ColumnQ_ItalicGray.xlsx");
    }
}