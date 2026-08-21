// Title: Export Excel to HTML with Bold & Italic Styling Using Aspose.Cells (C#)
// Description: This C# example creates a workbook, applies bold formatting to A1, italic to B1, and embeds HTML tags for combined bold‑italic text in C1. HtmlSaveOptions are configured to generate CSS classes (DisableCss = false) and to parse HTML tags inside cells, then the workbook is saved as StyledOutput.html with CSS‑based styling preserved.
// Keywords: Aspose.Cells HTML export C# | Excel to HTML with CSS | preserve bold formatting Aspose | preserve italic formatting Aspose | HtmlSaveOptions ParseHtmlTagInCell | generate CSS classes Aspose.Cells | styled HTML from workbook
// Common Searches: Aspose.Cells export Excel to HTML with CSS | keep bold and italic when saving Excel as HTML | parse HTML tags inside Excel cells Aspose | disable inline styles Aspose.Cells HTML output | C# example for HTML export with styling
// Developer Intent: Export an Excel workbook to HTML while retaining bold and italic formatting through generated CSS rather than inline styles.
// Use Cases: Publish a spreadsheet‑based report on a website with clean CSS styling. | Create email‑ready HTML from Excel that respects text emphasis without inline formatting. | Generate documentation pages from a workbook where headings are bold and notes are italic, using reusable CSS classes.
// AI Prompts: Show how to add underline formatting to a cell and have it appear in the exported HTML with CSS classes. | Provide code to reference an external CSS file instead of the automatically generated stylesheet when saving HTML with Aspose.Cells. | Explain how to customize the names of generated CSS classes for bold and italic styles via HtmlSaveOptions.

using System;
using Aspose.Cells;

// This C# example creates a workbook, applies bold formatting to A1, italic to B1, and embeds HTML tags for combined bold‑italic text in C1. HtmlSaveOptions are configured to generate CSS classes (DisableCss = false) and to parse HTML tags inside cells, then the workbook is saved as StyledOutput.html with CSS‑based styling preserved.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Cell A1: apply bold style
        Cell cellA1 = worksheet.Cells["A1"];
        cellA1.PutValue("Bold Text");
        Style styleA1 = cellA1.GetStyle();
        styleA1.Font.IsBold = true;
        cellA1.SetStyle(styleA1);

        // Cell B1: apply italic style
        Cell cellB1 = worksheet.Cells["B1"];
        cellB1.PutValue("Italic Text");
        Style styleB1 = cellB1.GetStyle();
        styleB1.Font.IsItalic = true;
        cellB1.SetStyle(styleB1);

        // Cell C1: use HTML tags to combine bold and italic
        worksheet.Cells["C1"].HtmlString = "<b>Bold</b> and <i>Italic</i>";

        // Configure HTML save options to use CSS (inline styles disabled)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.DisableCss = false;               // ensure CSS classes are generated
        htmlOptions.ParseHtmlTagInCell = true;         // parse HTML tags inside cells

        // Save the workbook as HTML with the specified options
        workbook.Save("StyledOutput.html", htmlOptions);
    }
}
