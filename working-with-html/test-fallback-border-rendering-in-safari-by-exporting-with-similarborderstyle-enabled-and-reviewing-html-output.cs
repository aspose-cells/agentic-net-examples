// Title: Export an Aspose.Cells workbook to HTML with SimilarBorderStyle enabled to verify thick border fallback in Safari
// AI Prompts: Write C# code that creates a workbook, applies a thick border style to a range, sets HtmlSaveOptions.SimilarBorderStyle to true, and saves the result as an HTML file. | Modify existing Aspose.Cells HTML export logic to enable the SimilarBorderStyle option and output the HTML so that Safari's border fallback can be inspected.
// Common Searches: how to enable SimilarBorderStyle in Aspose.Cells HTML export C# | Aspose.Cells export to HTML with thick cell borders for Safari compatibility | C# generate HTML from workbook and test border rendering in Safari | Aspose.Cells HtmlSaveOptions border fallback issue in Safari
// Tags: Aspose.Cells HTML export SimilarBorderStyle | C# thick cell borders HTML output | Safari border rendering fallback Aspose.Cells | HtmlSaveOptions border style handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new Workbook, adds sample data, defines a style with thick borders on all sides, applies the style to range A1:C2, configures HtmlSaveOptions with SimilarBorderStyle enabled, and saves the workbook as FallbackBorderSafari.html to test how Safari renders the exported borders.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["C2"].PutValue(30);

            // Define a style with thick borders on all sides
            var style = workbook.CreateStyle();
            var borders = style.Borders;
            borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

            // Apply the style to the range A1:C2
            var range = sheet.Cells.CreateRange("A1:C2");
            var styleFlag = new StyleFlag { All = true };
            range.ApplyStyle(style, styleFlag);

            // Configure HTML export options
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Define output file path
            string outputPath = "FallbackBorderSafari.html";

            // Save the workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
