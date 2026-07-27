// Title: Aspose.Cells .NET – Enable ExportSimilarBorderStyle in HtmlSaveOptions for fallback borders
// Description: C# example that creates a workbook, applies medium borders, sets HtmlSaveOptions.ExportSimilarBorderStyle to true, and saves as HTML so browsers lacking native border support receive equivalent styling.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportSimilarBorderStyle | C# | fallback borders | HTML export | Excel to HTML | border compatibility | Aspose.Cells .NET | similar border style
// Common Searches: Aspose.Cells ExportSimilarBorderStyle example | How to enable similar border style in HtmlSaveOptions C# | HTML export with border fallback Aspose.Cells | Save workbook as HTML with borders Aspose.Cells .NET | HtmlSaveOptions ExportSimilarBorderStyle property usage
// Developer Intent: Activate ExportSimilarBorderStyle so the generated HTML contains equivalent border CSS for browsers that cannot render the original Excel border formatting.
// Use Cases: Web dashboards that need exact Excel‑style borders across legacy browsers. | Automated report pipelines where HTML output must preserve cell borders. | Embedding styled worksheets in web pages with varying CSS border support. | Creating printable HTML versions of spreadsheets with consistent border appearance.
// AI Prompts: Generate C# code using Aspose.Cells to apply medium borders and export to HTML with ExportSimilarBorderStyle enabled. | Explain how ExportSimilarBorderStyle changes the HTML markup produced by Aspose.Cells and why it helps unsupported browsers. | Provide a step‑by‑step guide to configure HtmlSaveOptions for fallback borders in Aspose.Cells .NET.

using System;
using Aspose.Cells;

// C# example that creates a workbook, applies medium borders, sets HtmlSaveOptions.ExportSimilarBorderStyle to true, and saves as HTML so browsers lacking native border support receive equivalent styling.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to a cell
        worksheet.Cells["A1"].PutValue("Sample Data");

        // Create a style with medium borders on all sides
        Style borderStyle = workbook.CreateStyle();
        borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Medium;
        borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Medium;
        borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Medium;
        borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Medium;

        // Apply the border style to the cell
        worksheet.Cells["A1"].SetStyle(borderStyle);

        // Create HTML save options and enable fallback similar border style
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportSimilarBorderStyle = true
        };

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
