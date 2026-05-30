using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExportSimilarBorderStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Helper method to create a style with a specific border type
            Style CreateBorderStyle(Workbook wb, CellBorderType borderType, Color borderColor)
            {
                Style style = wb.CreateStyle();
                style.Borders[BorderType.TopBorder].LineStyle = borderType;
                style.Borders[BorderType.BottomBorder].LineStyle = borderType;
                style.Borders[BorderType.LeftBorder].LineStyle = borderType;
                style.Borders[BorderType.RightBorder].LineStyle = borderType;

                style.Borders[BorderType.TopBorder].Color = borderColor;
                style.Borders[BorderType.BottomBorder].Color = borderColor;
                style.Borders[BorderType.LeftBorder].Color = borderColor;
                style.Borders[BorderType.RightBorder].Color = borderColor;

                return style;
            }

            // Cell A1 – uses a border style that older browsers may not support (Double)
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue("Double Border");
            cellA1.SetStyle(CreateBorderStyle(workbook, CellBorderType.Double, Color.Blue));

            // Cell B1 – uses a medium border (also may be unsupported)
            Cell cellB1 = sheet.Cells["B1"];
            cellB1.PutValue("Medium Border");
            cellB1.SetStyle(CreateBorderStyle(workbook, CellBorderType.Medium, Color.Green));

            // Cell C1 – uses a simple thin border (widely supported)
            Cell cellC1 = sheet.Cells["C1"];
            cellC1.PutValue("Thin Border");
            cellC1.SetStyle(CreateBorderStyle(workbook, CellBorderType.Thin, Color.Red));

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export a similar border style when the original style is not supported
                ExportSimilarBorderStyle = true,

                // Enable IE compatibility to simulate legacy browser rendering
                IsIECompatible = true,

                // Optional: keep borders collapsed for a cleaner table layout
                IsBorderCollapsed = true
            };

            // Save the workbook as HTML
            string outputPath = "SimilarBorderStyle.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with ExportSimilarBorderStyle enabled.");
        }
    }
}