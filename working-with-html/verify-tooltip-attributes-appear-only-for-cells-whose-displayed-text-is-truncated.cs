// Title: Aspose.Cells .NET – Verify HTML tooltip appears only for truncated cells
// Description: C# example that creates a workbook, sets a narrow column width to force truncation, enables AddTooltipText in HtmlSaveOptions, saves to HTML, reads the output, counts title attributes and confirms that a tooltip is generated solely for the cell whose text is cut off.
// Keywords: Aspose.Cells HTML tooltip | AddTooltipText option | truncated cell detection | HTML export column width | title attribute verification | C# Aspose.Cells example
// Common Searches: Aspose.Cells add tooltip only for overflow text | HTML export tooltip for truncated cells Aspose | how to count title attributes in Aspose.Cells HTML output | verify tooltip generation with HtmlSaveOptions | C# test for truncated cell tooltips
// Developer Intent: Confirm that tooltips are added exclusively to cells whose displayed text is truncated after HTML export.
// Use Cases: Generate HTML reports where long cell values show full content on hover while short values remain tooltip‑free. | Automate regression tests that validate tooltip count matches the number of overflow cells in CI pipelines. | Improve user experience in web‑based spreadsheets by conditionally adding hover text based on column width.
// AI Prompts: Create C# code using Aspose.Cells to export a workbook to HTML with AddTooltipText enabled and verify that only truncated cells contain a title attribute. | Write an NUnit test that parses the generated HTML and asserts the number of title attributes equals the number of cells with overflow text. | Explain how SaveFormat.Html and HtmlSaveOptions properties (AddTooltipText, FormatDataIgnoreColumnWidth) affect tooltip generation for truncated cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTooltipVerification
{
    // C# example that creates a workbook, sets a narrow column width to force truncation, enables AddTooltipText in HtmlSaveOptions, saves to HTML, reads the output, counts title attributes and confirms that a tooltip is generated solely for the cell whose text is cut off.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Cell A1: short text (will fit)
            cells["A1"].PutValue("Short");

            // Cell A2: long text (will be truncated)
            cells["A2"].PutValue("This is a very long text that will not fit into the column width and should be truncated");

            // Set a narrow column width so that A2 gets truncated
            cells.SetColumnWidth(0, 10); // Column A width

            // Configure HTML save options to add tooltip text when data is truncated
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.AddTooltipText = true;               // Enable tooltip generation
            saveOptions.FormatDataIgnoreColumnWidth = false; // Ensure truncation occurs (default)

            // Define output HTML file path
            string htmlPath = "TooltipVerification.html";

            // Save the workbook as HTML
            workbook.Save(htmlPath, saveOptions);

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Simple verification:
            // Count occurrences of the title attribute (tooltip) in the HTML.
            // Expect exactly one tooltip for the truncated cell (A2) and none for the short cell (A1).
            int tooltipCount = 0;
            int index = 0;
            while ((index = htmlContent.IndexOf("title=", index, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                tooltipCount++;
                index += 6; // Move past "title="
            }

            // Output verification result
            if (tooltipCount == 1)
            {
                Console.WriteLine("Verification passed: Tooltip appears only for the truncated cell.");
            }
            else
            {
                Console.WriteLine($"Verification failed: Expected 1 tooltip, found {tooltipCount}.");
            }
        }
    }
}
