// Title: Generate HTML from an Excel workbook in C# with full cell text shown in a tooltip while preserving column overflow using Aspose.Cells
// AI Prompts: Write C# code that creates a workbook, narrows a column, inserts a long string into a cell, and saves the sheet as HTML with HtmlSaveOptions.AddTooltipText enabled so the tooltip displays the entire cell value and the overflow layout stays unchanged. | Show how to set up Aspose.Cells HtmlSaveOptions to keep cell overflow and activate tooltip text for long cell contents when exporting a worksheet to HTML in .NET.
// Common Searches: Aspose.Cells C# export worksheet to HTML with tooltip for long cell values | How to preserve column overflow and show full cell text in HTML using Aspose.Cells | Enable AddTooltipText in HtmlSaveOptions to display cell content on mouseover | Keep Excel cell overflow when converting to HTML with Aspose.Cells .NET | Show full cell content in HTML tooltip Aspose.Cells example
// Tags: Aspose.Cells HtmlSaveOptions AddTooltipText | HTML export preserve cell overflow | C# tooltip for full cell content | Excel to HTML conversion with tooltip | HtmlCrossType.Default overflow handling

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, narrows the first column, places a long text string in cell A1, enables AddTooltipText in HtmlSaveOptions, and saves the workbook as HTML. The resulting HTML retains the original overflow appearance while displaying the complete cell value in a mouseover tooltip.
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

            // Set a narrow column width to force overflow
            sheet.Cells.SetColumnWidth(0, 10);

            // Put a long text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("This is a very long text that will overflow the cell width and should be shown in a tooltip.");

            // Optional: set alignment to demonstrate overflow
            Style style = cell.GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Left;
            cell.SetStyle(style);

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Enable tooltip with full cell content
            htmlOptions.AddTooltipText = true;

            // Define output file path
            string outputPath = "output.html";

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
