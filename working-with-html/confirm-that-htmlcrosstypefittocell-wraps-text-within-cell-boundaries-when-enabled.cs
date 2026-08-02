// Title: Verify HtmlCrossType.FitToCell Wraps Text Inside Cell When Exporting to HTML with Aspose.Cells for .NET
// Description: This C# sample creates a workbook, puts a long string into cell A1, enables text wrapping, narrows the column, sets HtmlSaveOptions.HtmlCrossStringType to FitToCell, and saves the workbook as HTML. The generated HTML shows the text confined to the cell without spilling into adjacent columns.
// Keywords: Aspose.Cells | HtmlCrossType.FitToCell | HTML export | text wrapping | .NET | cell overflow | HtmlSaveOptions | Excel to HTML | wrap long text | cell width
// Common Searches: Aspose.Cells HtmlCrossType FitToCell example | How to prevent text overflow in HTML export using Aspose.Cells | Wrap long cell text when saving workbook as HTML .NET | FitToCell behavior Aspose.Cells HTML | C# export Excel to HTML with text wrapping
// Developer Intent: Confirm that setting HtmlCrossStringType to FitToCell forces cell text to wrap and stay within the cell limits in the generated HTML.
// Use Cases: Generate web‑friendly spreadsheet reports where column widths are fixed and text must stay inside cells. | Create printable HTML views of spreadsheets that preserve the original layout without overflow. | Embed Excel data in web pages or dashboards while ensuring long strings wrap correctly. | Automate HTML export for reporting tools that require consistent cell dimensions.
// AI Prompts: Write a C# unit test that loads the saved HTML file and asserts that the CSS width of cell A1 matches the column width and that the text does not overflow. | Provide a PowerShell script to parse the generated HTML and verify that the style attribute for cell A1 includes text‑wrap and width constraints after using FitToCell. | Explain how to combine HtmlCrossType.FitToCell with AutoFitRows to ensure all wrapped text is fully visible in the HTML output.

using System;
using Aspose.Cells;

// This C# sample creates a workbook, puts a long string into cell A1, enables text wrapping, narrows the column, sets HtmlSaveOptions.HtmlCrossStringType to FitToCell, and saves the workbook as HTML. The generated HTML shows the text confined to the cell without spilling into adjacent columns.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a long text that would normally overflow the cell width
        worksheet.Cells["A1"].PutValue(
            "This is a very long text that should be wrapped and stay inside the cell when HtmlCrossType.FitToCell is used.");

        // Enable text wrapping for the cell
        Style style = worksheet.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(style);

        // Reduce column width to force the text to exceed the cell boundaries
        worksheet.Cells.SetColumnWidth(0, 10);

        // Configure HTML save options to use FitToCell behavior
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.HtmlCrossStringType = HtmlCrossType.FitToCell;

        // Save the workbook as HTML; the resulting file will show the text confined within the cell
        workbook.Save("FitToCellDemo.html", htmlOptions);
    }
}
