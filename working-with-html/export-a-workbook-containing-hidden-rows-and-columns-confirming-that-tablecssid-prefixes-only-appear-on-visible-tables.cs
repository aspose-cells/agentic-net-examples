// Title: Export Excel to HTML without Hidden Rows/Columns and Apply Custom TableCssId – Aspose.Cells C# Example
// Description: Demonstrates how to create a workbook, hide specific rows and columns, configure HtmlSaveOptions to remove hidden elements, set a custom TableCssId, save as HTML, and verify that the CSS identifier appears only on the visible table.
// Keywords: Aspose.Cells HTML export | remove hidden rows C# | remove hidden columns Aspose | TableCssId custom prefix | verify TableCssId visibility | HtmlSaveOptions HiddenRowDisplayType | HtmlHiddenColDisplayType | C# Excel to HTML example
// Common Searches: Aspose.Cells export Excel to HTML without hidden rows | How to hide rows and columns when saving as HTML in .NET | Set custom TableCssId for HTML tables using Aspose.Cells | Remove hidden columns from HTML output with Aspose | Validate TableCssId appears only on visible tables
// Developer Intent: Generate an HTML file from an Excel workbook that excludes hidden rows/columns and applies a custom TableCssId solely to the visible table.
// Use Cases: Create clean web‑ready reports where confidential rows or columns are hidden in the source workbook. | Apply a consistent CSS class to exported tables for styling while ensuring hidden data is not rendered. | Automate validation that the custom TableCssId occurs the expected number of times after export.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, removing hidden rows and columns and assigning a custom TableCssId. | Explain the effect of HtmlHiddenRowDisplayType.Remove and HtmlHiddenColDisplayType.Remove on the generated HTML. | Show how to programmatically count occurrences of a specific TableCssId in the exported HTML to confirm it only belongs to visible tables.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, hide specific rows and columns, configure HtmlSaveOptions to remove hidden elements, set a custom TableCssId, save as HTML, and verify that the CSS identifier appears only on the visible table.
class ExportHiddenRowsColumnsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Sheet1";

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");
        worksheet.Cells["A2"].PutValue("Data1");
        worksheet.Cells["B2"].PutValue("Data2");
        worksheet.Cells["C2"].PutValue("Data3");

        // Hide row 2 (index 1) and column B (index 1)
        worksheet.Cells.HideRow(1);
        worksheet.Cells.HideColumn(1);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Set a custom prefix for CSS classes generated for tables
            TableCssId = "myTable",
            // Remove hidden rows/columns from the generated HTML
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove,
            HiddenColDisplayType = HtmlHiddenColDisplayType.Remove
        };

        // Define output file path
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.html");

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, saveOptions);

        // Load the generated HTML to verify that the TableCssId appears only for visible tables
        string htmlContent = File.ReadAllText(outputPath);
        int occurrenceCount = (htmlContent.Length - htmlContent.Replace("myTable", "").Length) / "myTable".Length;
        Console.WriteLine($"TableCssId occurrences in the HTML: {occurrenceCount}");
        // The count should correspond to the visible table only (hidden rows/columns are removed)
    }
}
