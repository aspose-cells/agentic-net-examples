using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with actual file path if needed)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Define a style that mirrors the formatting from the reference XLSX
            // ------------------------------------------------------------
            Style style = workbook.CreateStyle();

            // Font settings
            style.Font.Name = "Calibri";
            style.Font.Size = 11;
            style.Font.IsBold = true;
            style.Font.Color = Color.White;

            // Cell shading (background)
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.DarkBlue;

            // Borders
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].Color = Color.Black;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;
            style.Borders[BorderType.LeftBorder].Color = Color.Black;
            style.Borders[BorderType.RightBorder].Color = Color.Black;

            // Alignment
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.IsTextWrapped = true;

            // ------------------------------------------------------------
            // Apply the style to specific ranges
            // ------------------------------------------------------------

            // Example range 1: A1:C1 (header row)
            AsposeRange headerRange = cells.CreateRange("A1", "C1");
            headerRange.SetStyle(style); // Apply full style to the range

            // Example range 2: A2:C10 (data area) – only apply borders and alignment
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            dataStyle.HorizontalAlignment = TextAlignmentType.Left;
            dataStyle.VerticalAlignment = TextAlignmentType.Center;

            StyleFlag dataFlag = new StyleFlag
            {
                Borders = true,
                HorizontalAlignment = true,
                VerticalAlignment = true
            };

            AsposeRange dataRange = cells.CreateRange("A2", "C10");
            dataRange.ApplyStyle(dataStyle, dataFlag); // Apply selective formatting

            // Example range 3: D1:D10 – apply only background shading
            Style shadeStyle = workbook.CreateStyle();
            shadeStyle.Pattern = BackgroundType.Solid;
            shadeStyle.ForegroundColor = Color.LightGray;

            StyleFlag shadeFlag = new StyleFlag
            {
                CellShading = true
            };

            AsposeRange shadeRange = cells.CreateRange("D1", "D10");
            shadeRange.ApplyStyle(shadeStyle, shadeFlag);

            // ------------------------------------------------------------
            // Save the modified workbook
            // ------------------------------------------------------------
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}