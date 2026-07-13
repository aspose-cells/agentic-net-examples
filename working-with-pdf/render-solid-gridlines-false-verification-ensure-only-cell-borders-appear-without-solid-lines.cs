using System;
using System.Drawing;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – hides default gridlines and shows only cell borders
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide the default worksheet gridlines
        worksheet.IsGridlinesVisible = false;

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Item1");
        worksheet.Cells["B2"].PutValue(100);

        // Define a style with thin black borders on all sides
        Style borderStyle = workbook.CreateStyle();
        borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
        borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

        // Apply the border style to the range containing the data
        StyleFlag flag = new StyleFlag { All = true };
        worksheet.Cells.CreateRange("A1:B2").ApplyStyle(borderStyle, flag);

        // Save the workbook
        workbook.Save("BordersOnly.xlsx");
    }
}