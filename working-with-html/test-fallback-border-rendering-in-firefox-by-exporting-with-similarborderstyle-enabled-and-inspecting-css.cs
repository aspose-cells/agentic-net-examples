// Title: Export Double Borders to HTML with ExportSimilarBorderStyle in Aspose.Cells for .NET (Firefox CSS Check)
// Description: Creates a workbook, applies a blue double‑line border to cell A1, enables ExportSimilarBorderStyle and disables border collapse in HtmlSaveOptions, saves the file as HTML, and guides the developer to open it in Firefox to inspect the generated fallback CSS for unsupported double borders.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportSimilarBorderStyle | double border | CSS fallback | Firefox | C# | HTML export | border rendering | unsupported border types
// Common Searches: Aspose.Cells ExportSimilarBorderStyle example | how to export double border to HTML | Firefox CSS for double cell border Aspose | HTML border rendering Aspose.Cells .NET | fallback border style Aspose.Cells
// Developer Intent: Generate an HTML file that uses ExportSimilarBorderStyle to provide fallback CSS for double borders and verify the output in Firefox.
// Use Cases: Export workbooks with double or other unsupported borders to HTML for cross‑browser compatibility. | Debug border rendering by disabling border collapse to view individual CSS rules. | Automate QA testing of HTML output in Firefox to ensure correct fallback styling.
// AI Prompts: Show how to also export the workbook as PDF while preserving double borders. | Provide a script that opens the saved HTML in Firefox and extracts the computed CSS for cell A1. | Explain the internal mapping of CellBorderType.Double to CSS properties when ExportSimilarBorderStyle is enabled.

using System;
using Aspose.Cells;
using System.Drawing;

// Creates a workbook, applies a blue double‑line border to cell A1, enables ExportSimilarBorderStyle and disables border collapse in HtmlSaveOptions, saves the file as HTML, and guides the developer to open it in Firefox to inspect the generated fallback CSS for unsupported double borders.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define a style with a border type that many browsers do not support (Double)
        Style doubleBorderStyle = workbook.CreateStyle();
        doubleBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
        doubleBorderStyle.Borders[BorderType.TopBorder].Color = Color.Blue;
        doubleBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
        doubleBorderStyle.Borders[BorderType.BottomBorder].Color = Color.Blue;
        doubleBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
        doubleBorderStyle.Borders[BorderType.LeftBorder].Color = Color.Blue;
        doubleBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
        doubleBorderStyle.Borders[BorderType.RightBorder].Color = Color.Blue;

        // Apply the style to a cell with some content
        worksheet.Cells["A1"].PutValue("Double Border");
        worksheet.Cells["A1"].SetStyle(doubleBorderStyle);

        // Configure HTML save options to export similar border style (fallback) and keep borders separate
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportSimilarBorderStyle = true, // Enable fallback rendering for unsupported borders
            IsBorderCollapsed = false        // Keep table borders uncollapsed for clearer CSS inspection
        };

        // Save the workbook as HTML
        string outputFile = "FallbackBorder.html";
        workbook.Save(outputFile, htmlOptions);

        Console.WriteLine($"HTML file saved to '{outputFile}'. Open it in Firefox and inspect the generated CSS for fallback border styles.");
    }
}
