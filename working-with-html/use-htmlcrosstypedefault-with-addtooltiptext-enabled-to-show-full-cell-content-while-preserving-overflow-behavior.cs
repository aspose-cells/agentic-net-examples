// Title: Export Excel to HTML with full cell tooltip using HtmlCrossType.Default and AddTooltipText in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet to HTML with tooltip showing complete cell text | How to keep cell overflow and add tooltip on HTML export using HtmlCrossType.Default | Enable AddTooltipText in HtmlSaveOptions to display long cell content on hover | C# Aspose.Cells HTML export preserve overflow and show full text in tooltip | Set HtmlCrossType.Default and tooltip for long text cells in Aspose.Cells HTML output
// Tags: Aspose.Cells HtmlSaveOptions AddTooltipText | HtmlCrossType.Default overflow handling | C# generate HTML from Excel with tooltip | preserve cell overflow in HTML export | full cell text tooltip on hover

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, inserts a long string into cell A1, narrows the column to force overflow, disables text wrapping, enables AddTooltipText in HtmlSaveOptions, and saves the file as HTML. The resulting HTML keeps the overflow appearance and shows the entire cell content in a tooltip when the user hovers over the cell.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text into a cell to demonstrate overflow
            sheet.Cells["A1"].PutValue("This is a very long text that will overflow the cell width, and we want the full content to appear in a tooltip when hovering over the cell.");

            // Set a small column width to force overflow
            sheet.Cells.SetColumnWidth(0, 10); // Column A

            // Ensure text does not wrap so overflow behavior is preserved
            Style style = workbook.CreateStyle();
            style.IsTextWrapped = false;
            sheet.Cells["A1"].SetStyle(style);

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.AddTooltipText = true; // Enable tooltip showing full cell content

            // Define output file path
            string outputPath = "output.html";

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
