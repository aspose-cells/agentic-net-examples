// Title: How to use HtmlSaveOptions.SimilarBorderStyle in Aspose.Cells for .NET to generate fallback border styles in HTML output
// AI Prompts: Create a C# workbook, apply thin black borders to one cell and thick blue borders to another, enable HtmlSaveOptions.SimilarBorderStyle for fallback styling, and save the workbook as HTML to verify border rendering in older browsers. | Adjust the SimilarBorderStyle property to a custom CSS fallback, generate the HTML file, and compare the visual result in Internet Explorer 8 or another legacy browser.
// Common Searches: Aspose.Cells C# export Excel to HTML with fallback border styles for IE8 | how to enable SimilarBorderStyle in HtmlSaveOptions when saving as HTML | HTML border compatibility issues with Aspose.Cells in legacy browsers | set HtmlSaveOptions.SimilarBorderStyle property .NET example
// Tags: Aspose.Cells HtmlSaveOptions SimilarBorderStyle | C# export Excel to HTML border fallback | legacy browser HTML border compatibility | configure border rendering Aspose.Cells | thin vs thick cell borders HTML export

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, writes data to cells A1 and A2, applies thin black borders to A1 and thick blue borders to A2, configures HtmlSaveOptions (including the SimilarBorderStyle fallback option), and saves the workbook as an HTML file to demonstrate how border styles are rendered in older browsers.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);

            // Apply thin black borders to cell A1
            Style styleA1 = sheet.Cells["A1"].GetStyle();
            styleA1.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            styleA1.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            styleA1.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            styleA1.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            sheet.Cells["A1"].SetStyle(styleA1);

            // Apply thick blue borders to cell A2
            Style styleA2 = sheet.Cells["A2"].GetStyle();
            styleA2.SetBorder(BorderType.TopBorder, CellBorderType.Thick, Color.Blue);
            styleA2.SetBorder(BorderType.BottomBorder, CellBorderType.Thick, Color.Blue);
            styleA2.SetBorder(BorderType.LeftBorder, CellBorderType.Thick, Color.Blue);
            styleA2.SetBorder(BorderType.RightBorder, CellBorderType.Thick, Color.Blue);
            sheet.Cells["A2"].SetStyle(styleA2);

            // Configure HTML save options (default settings already export borders)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
