using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsSafariBorderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Helper method to create a style with a specific border type
            Style CreateBorderStyle(CellBorderType borderType, Color color)
            {
                Style style = workbook.CreateStyle();
                style.Borders[BorderType.TopBorder].LineStyle = borderType;
                style.Borders[BorderType.BottomBorder].LineStyle = borderType;
                style.Borders[BorderType.LeftBorder].LineStyle = borderType;
                style.Borders[BorderType.RightBorder].LineStyle = borderType;
                style.Borders[BorderType.TopBorder].Color = color;
                style.Borders[BorderType.BottomBorder].Color = color;
                style.Borders[BorderType.LeftBorder].Color = color;
                style.Borders[BorderType.RightBorder].Color = color;
                return style;
            }

            // Cell A1 – supported thin border
            sheet.Cells["A1"].PutValue("Thin Border");
            sheet.Cells["A1"].SetStyle(CreateBorderStyle(CellBorderType.Thin, Color.Black));

            // Cell B1 – supported medium border
            sheet.Cells["B1"].PutValue("Medium Border");
            sheet.Cells["B1"].SetStyle(CreateBorderStyle(CellBorderType.Medium, Color.Blue));

            // Cell C1 – border type not widely supported (Double)
            sheet.Cells["C1"].PutValue("Double Border (fallback test)");
            sheet.Cells["C1"].SetStyle(CreateBorderStyle(CellBorderType.Double, Color.Red));

            // Configure HTML save options with ExportSimilarBorderStyle enabled
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportSimilarBorderStyle = true // Enable fallback rendering for unsupported borders
            };

            // Save the workbook as HTML
            string outputPath = "SafariBorderFallback.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}'. Open it in Safari to verify fallback border rendering.");
        }
    }
}