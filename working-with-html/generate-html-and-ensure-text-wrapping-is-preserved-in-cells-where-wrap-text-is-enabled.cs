// Title: C# – Export Excel to HTML with Preserved Text Wrapping Using Aspose.Cells
// Description: Demonstrates how to create a workbook, insert a long string, enable text wrapping via a style and StyleFlag, set column width, configure HtmlSaveOptions (HideOverflowWrappedText = false, FormatDataIgnoreColumnWidth = false), and save the file as HTML so that wrapped text appears correctly in the browser.
// Keywords: Aspose.Cells HTML export C# | preserve text wrapping Excel to HTML | HtmlSaveOptions HideOverflowWrappedText | wrap text style Aspose.Cells | C# Excel to HTML conversion | column width HTML wrap Aspose | Aspose.Cells .NET HTML output | export wrapped cells to HTML | Excel cell wrap HTML rendering
// Common Searches: how to keep text wrapping when saving Excel as HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions show wrapped text in HTML | C# export cell with wrap text to HTML | prevent overflow hidden for wrapped cells Aspose HTML export | preserve Excel cell line breaks in HTML output | set column width for HTML wrap Aspose.Cells
// Developer Intent: Generate an HTML file from an Excel workbook where cells that have text‑wrap enabled retain their line breaks and are fully visible in the rendered page.
// Use Cases: Web‑based reports that need multi‑line descriptions to display exactly as in Excel. | Exporting data sheets with comments or notes that rely on cell wrapping for readability. | Creating HTML invoices or catalogs where product details must wrap within table cells without truncation.
// AI Prompts: Show me how to add a custom CSS class to wrapped cells when exporting to HTML with Aspose.Cells. | Provide code to export multiple worksheets to a single HTML file while preserving text wrapping on all cells. | Explain how to adjust HtmlSaveOptions to control column width and prevent overflow for wrapped text in the generated HTML.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, insert a long string, enable text wrapping via a style and StyleFlag, set column width, configure HtmlSaveOptions (HideOverflowWrappedText = false, FormatDataIgnoreColumnWidth = false), and save the file as HTML so that wrapped text appears correctly in the browser.
class HtmlWrapDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert a long text into cell A1
        cells["A1"].PutValue("This is a very long text that should wrap inside the cell when exported to HTML. It demonstrates how wrap text is preserved.");

        // Create a style with text wrapping enabled
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;

        // Use a StyleFlag to apply only the wrap text setting
        StyleFlag flag = new StyleFlag();
        flag.WrapText = true;

        // Apply the wrapping style to cell A1
        cells["A1"].SetStyle(wrapStyle, flag);

        // Set column width to force the text to wrap
        cells.SetColumnWidth(0, 15); // width in characters

        // Configure HTML save options to keep wrapped text visible
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.HideOverflowWrappedText = false;          // do not hide overflow
        htmlOptions.FormatDataIgnoreColumnWidth = false;     // respect column width like Excel

        // Save the workbook as an HTML file
        workbook.Save("WrappedTextOutput.html", htmlOptions);
    }
}
