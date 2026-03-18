using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to demonstrate borders
        worksheet.Cells["A1"].PutValue("Unsupported Border");
        worksheet.Cells["B1"].PutValue("Normal Cell");

        // Define a style with a border type that browsers typically do not support (e.g., Double)
        Style borderStyle = workbook.CreateStyle();
        borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
        borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
        borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
        borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;

        // Apply the style to cell A1
        worksheet.Cells["A1"].SetStyle(borderStyle);

        // Configure HTML save options to approximate unsupported border styles
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportSimilarBorderStyle = true // Enables similar border approximation
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("BorderApproximation.html", htmlOptions);
    }
}