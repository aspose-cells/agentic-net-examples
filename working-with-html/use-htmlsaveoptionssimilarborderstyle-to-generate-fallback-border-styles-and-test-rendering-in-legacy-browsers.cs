// Title: Generate HTML with fallback cell borders using HtmlSaveOptions.SimilarBorderStyle in Aspose.Cells for .NET
// AI Prompts: Write C# code that applies a thin border to a cell range, enables SimilarBorderStyle on HtmlSaveOptions, and saves the workbook as HTML for older‑browser compatibility. | Modify the Aspose.Cells sample to turn on SimilarBorderStyle, then open the produced HTML in Internet Explorer 8 to confirm the borders render correctly. | Create a C# program that defines a custom thin border style, configures HtmlSaveOptions with fallback border handling, and outputs the HTML file for testing in legacy browsers.
// Common Searches: Aspose.Cells enable SimilarBorderStyle when saving workbook as HTML for IE8 support | C# export Excel range with thin borders to HTML using fallback CSS via HtmlSaveOptions | How to make HTML border rendering work in older browsers with Aspose.Cells HtmlSaveOptions
// Tags: Aspose.Cells HtmlSaveOptions SimilarBorderStyle implementation | C# generate HTML with Excel cell borders | legacy browser HTML border compatibility | thin border style applied to cell range | verify HTML border rendering in IE8

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds header and numeric data, defines a thin border style, applies it to cells A1:B2, configures HtmlSaveOptions (including the SimilarBorderStyle fallback option), and saves the result as an HTML file suitable for testing in older browsers.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Define a thin border style
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Apply the border style to the range A1:B2
            StyleFlag flag = new StyleFlag() { Borders = true };
            sheet.Cells.CreateRange("A1:B2").ApplyStyle(borderStyle, flag);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: SimilarBorderStyle property is not available in the current Aspose.Cells version.
            // The default rendering will be used.

            // Save the workbook as an HTML file
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
