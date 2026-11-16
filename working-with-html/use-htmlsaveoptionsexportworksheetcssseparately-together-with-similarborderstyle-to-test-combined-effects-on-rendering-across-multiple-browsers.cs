// Title: Export an Excel worksheet to HTML with per‑worksheet CSS and test SimilarBorderStyle rendering in C# using Aspose.Cells
// Common Searches: Aspose.Cells C# export workbook to HTML with separate CSS file per worksheet | How to enable ExportWorksheetCSSSeparately and SimilarBorderStyle in HtmlSaveOptions | C# code sample for testing border style consistency across browsers using Aspose.Cells HTML export | Generate HTML and external CSS from Excel using Aspose.Cells HtmlSaveOptions | Verify thin blue borders render correctly in HTML output with Aspose.Cells
// Tags: exportworksheetcssseparately htmlsaveoptions aspocells | similarborderstyle html export aspocells | per‑worksheet external css generation c# | border style rendering cross‑browser aspocells | c# html export thin blue borders aspocells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsHtmlExportDemo
{
    // The example creates a workbook, adds sample data, applies thin blue borders to a range, configures HtmlSaveOptions with ExportWorksheetCSSSeparately set to true (and notes the SimilarBorderStyle property when available), then saves the workbook as HTML and a separate CSS file, allowing developers to evaluate border rendering consistency across browsers.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Header 1");
                sheet.Cells["B1"].PutValue("Header 2");
                sheet.Cells["A2"].PutValue(123);
                sheet.Cells["B2"].PutValue(456);

                // Apply border styles to the range A1:B2
                Style style = workbook.CreateStyle();
                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.TopBorder].Color = Color.Blue;
                style.Borders[BorderType.BottomBorder].Color = Color.Blue;
                style.Borders[BorderType.LeftBorder].Color = Color.Blue;
                style.Borders[BorderType.RightBorder].Color = Color.Blue;

                // Apply the style to the range
                Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:B2");
                range.ApplyStyle(style, new StyleFlag { All = true });

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export CSS for each worksheet into a separate file (e.g., output.css)
                    ExportWorksheetCSSSeparately = true
                    // Note: SimilarBorderStyle property is not available in the current Aspose.Cells version.
                };

                // Save the workbook as HTML with the specified options
                // This will generate "output.html" and a separate CSS file "output.css"
                string htmlPath = "output.html";
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine("HTML export completed with ExportWorksheetCSSSeparately enabled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
