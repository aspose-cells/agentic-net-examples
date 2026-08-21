// Title: Aspose.Cells .NET: Export to HTML with overflow tooltip using HtmlCrossType.Default
// Description: Learn how to save a workbook as HTML with Aspose.Cells, forcing column overflow and showing the full cell text in a tooltip. The example uses HtmlSaveOptions, HtmlCrossType.Default, and AddTooltipText to preserve Excel‑like overflow behavior.
// Keywords: Aspose.Cells | .NET | C# | HTML export | HtmlSaveOptions | HtmlCrossType.Default | AddTooltipText | cell overflow tooltip | Excel style overflow | web spreadsheet
// Common Searches: Aspose.Cells HTML tooltip for overflow cells | Enable AddTooltipText in HtmlSaveOptions | HtmlCrossType.Default overflow behavior | Show full cell text on hover Aspose.Cells | Export Excel to HTML with tooltips C#
// Developer Intent: Export a workbook to HTML where cells that exceed column width keep the overflow display and reveal the complete content in a tooltip on mouse‑over.
// Use Cases: HTML reports with long descriptions that stay compact but are fully readable on hover. | Web‑based spreadsheets that mimic Excel overflow while providing accessibility via tooltips. | Dashboard tables where space is limited yet users need instant access to full cell values.
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook to HTML with HtmlCrossType.Default and AddTooltipText enabled for all cells. | Explain how HtmlCrossType.Default differs from other cross types and how AddTooltipText affects overflow rendering. | Suggest additional HtmlSaveOptions settings to style tooltips or control their appearance in the generated HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlTooltipDemo
{
    // Learn how to save a workbook as HTML with Aspose.Cells, forcing column overflow and showing the full cell text in a tooltip. The example uses HtmlSaveOptions, HtmlCrossType.Default, and AddTooltipText to preserve Excel‑like overflow behavior.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text into a cell that will exceed the column width
            sheet.Cells["A1"].PutValue("This is a very long text that will overflow the cell width and we want to see the full content in a tooltip.");

            // Set a narrow column width to force overflow
            sheet.Cells.SetColumnWidth(0, 8); // width in characters

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Use the default HTML cross type (behaves like Excel)
            htmlOptions.HtmlCrossStringType = HtmlCrossType.Default;
            // Enable tooltip text for cells where content cannot be fully displayed
            htmlOptions.AddTooltipText = true;

            // Save the workbook as HTML
            string outputPath = "OverflowWithTooltip.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
