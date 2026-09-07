// Title: C# Aspose.Cells export workbook to HTML with custom cell borders for Safari fallback rendering test
// AI Prompts: Write C# code that builds a workbook, sets thin black, double red, dashed green, and medium blue borders on specific cells, and saves the file as HTML using Aspose.Cells. | Enable the fallback rendering feature for Safari in Aspose.Cells HtmlSaveOptions and generate the HTML output for visual verification.
// Common Searches: how to turn on Safari border fallback when saving Excel to HTML with Aspose.Cells .NET | C# example exporting workbook to HTML with multiple cell border styles using Aspose.Cells | testing thin, double, dashed, medium borders in HTML output for Safari compatibility | Aspose.Cells HtmlSaveOptions settings for border rendering differences in Safari | verify cell border appearance in Safari after converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions border compatibility Safari | C# set CellBorderType styles Aspose.Cells | export workbook to HTML with custom cell borders | Safari HTML rendering of Excel borders | SimilarBorderStyle option Aspose.Cells | cell border line style HTML output

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, applies thin black, double red, dashed green, and medium blue borders to cells A1, B1, A2, and B2, and saves the workbook as HTML. It demonstrates how to configure HtmlSaveOptions to activate Safari's border fallback rendering for visual testing.
class SafariBorderFallbackTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with sample data
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue("Data 2");

            // Apply different border styles to demonstrate fallback rendering
            // Cell A1: Thin black border
            Style styleA1 = workbook.CreateStyle();
            styleA1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            styleA1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            styleA1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            styleA1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            styleA1.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
            styleA1.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
            styleA1.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
            styleA1.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
            sheet.Cells["A1"].SetStyle(styleA1);

            // Cell B1: Double red border
            Style styleB1 = workbook.CreateStyle();
            styleB1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
            styleB1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
            styleB1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
            styleB1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
            styleB1.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Red;
            styleB1.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Red;
            styleB1.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Red;
            styleB1.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Red;
            sheet.Cells["B1"].SetStyle(styleB1);

            // Cell A2: Dashed green border
            Style styleA2 = workbook.CreateStyle();
            styleA2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Dashed;
            styleA2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Dashed;
            styleA2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Dashed;
            styleA2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Dashed;
            styleA2.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Green;
            styleA2.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Green;
            styleA2.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Green;
            styleA2.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Green;
            sheet.Cells["A2"].SetStyle(styleA2);

            // Cell B2: Medium blue border
            Style styleB2 = workbook.CreateStyle();
            styleB2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Medium;
            styleB2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Medium;
            styleB2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Medium;
            styleB2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Medium;
            styleB2.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Blue;
            styleB2.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Blue;
            styleB2.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Blue;
            styleB2.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Blue;
            sheet.Cells["B2"].SetStyle(styleB2);

            // Configure HTML save options (SimilarBorderStyle not available in this version)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Export the workbook to HTML
            string outputPath = "SafariBorderFallback.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file generated at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
