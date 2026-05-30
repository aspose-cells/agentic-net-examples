using System;
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

            // Add sample data to demonstrate borders
            sheet.Cells["A1"].PutValue("Cell with thick border");
            sheet.Cells["B1"].PutValue("Cell with thin border");

            // Apply a thick border style to A1
            Style thickStyle = workbook.CreateStyle();
            thickStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            thickStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            thickStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            thickStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            sheet.Cells["A1"].SetStyle(thickStyle);

            // Apply a thin border style to B1
            Style thinStyle = workbook.CreateStyle();
            thinStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            thinStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            thinStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            thinStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            sheet.Cells["B1"].SetStyle(thinStyle);

            // Create HTML save options and enable similar border style fallback
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportSimilarBorderStyle = true // Fallback for browsers that don't support certain border styles
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportSimilarBorderStyle.html", htmlOptions);

            Console.WriteLine("Workbook saved as HTML with ExportSimilarBorderStyle enabled.");
        }
    }
}