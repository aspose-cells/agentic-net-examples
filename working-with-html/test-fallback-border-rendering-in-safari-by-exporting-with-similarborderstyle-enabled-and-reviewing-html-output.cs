// Title: Test Safari Border Fallback with Aspose.Cells HTML Export (ExportSimilarBorderStyle) – C# Example
// Description: Creates a workbook, applies thin, double and medium dash‑dot borders to cells, enables HtmlSaveOptions.ExportSimilarBorderStyle, saves as HTML, and shows how Safari substitutes unsupported border styles with similar ones.
// Keywords: Aspose.Cells | C# HTML export | ExportSimilarBorderStyle | Safari border fallback | CellBorderType | HtmlSaveOptions | .NET cross‑browser rendering | Excel to HTML conversion | border style mapping | HTML report generation
// Common Searches: Aspose.Cells ExportSimilarBorderStyle Safari example | how to test border rendering in Safari with Aspose.Cells | unsupported Excel border types in Safari HTML output | C# code to export Excel to HTML with fallback borders | Aspose.Cells HTML export border compatibility
// Developer Intent: Export an Excel workbook to HTML with ExportSimilarBorderStyle enabled to verify how Safari renders borders that are not natively supported.
// Use Cases: Produce HTML reports that retain visual border consistency across Chrome, Firefox, and Safari. | Automate visual regression tests for border appearance in Safari versus other browsers. | Generate printable invoices or dashboards where double or dash‑dot borders are automatically replaced with Safari‑friendly styles.
// AI Prompts: Guide me through opening SafariBorderFallback.html in Safari and describing the rendered borders for each cell. | Show how to modify the sample to log cells whose border style was altered by ExportSimilarBorderStyle. | Explain the CSS mapping Aspose.Cells uses for unsupported CellBorderType values when ExportSimilarBorderStyle is true.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsSafariBorderFallbackDemo
{
    // Creates a workbook, applies thin, double and medium dash‑dot borders to cells, enables HtmlSaveOptions.ExportSimilarBorderStyle, saves as HTML, and shows how Safari substitutes unsupported border styles with similar ones.
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
                // Apply the same border style to all four sides
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

            // Cell A1 – supported thin border (baseline)
            sheet.Cells["A1"].PutValue("Thin Border");
            sheet.Cells["A1"].SetStyle(CreateBorderStyle(CellBorderType.Thin, Color.Black));

            // Cell B1 – double border (not supported by some browsers like Safari)
            sheet.Cells["B1"].PutValue("Double Border");
            sheet.Cells["B1"].SetStyle(CreateBorderStyle(CellBorderType.Double, Color.Blue));

            // Cell C1 – medium dash dot border (also potentially unsupported)
            sheet.Cells["C1"].PutValue("MediumDashDot Border");
            sheet.Cells["C1"].SetStyle(CreateBorderStyle(CellBorderType.MediumDashDot, Color.Green));

            // Configure HTML save options with ExportSimilarBorderStyle enabled
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportSimilarBorderStyle = true // Fallback to a similar style when original is unsupported
            };

            // Save the workbook as HTML
            string outputPath = "SafariBorderFallback.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}'. Open it in Safari to verify fallback border rendering.");
        }
    }
}
