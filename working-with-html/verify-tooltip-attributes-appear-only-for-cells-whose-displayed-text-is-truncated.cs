// Title: Aspose.Cells for .NET – Verify HTML tooltips only for truncated cell text
// Description: This C# example creates a workbook with two cells, narrows the first column to force overflow, enables AddTooltipText in HtmlSaveOptions, saves to HTML, and programmatically checks that a <title> attribute is generated for the overflowed cell (A1) while the fitting cell (B1) remains tooltip‑free.
// Keywords: Aspose.Cells HTML tooltip | AddTooltipText .NET | truncated cell text tooltip | verify title attribute Aspose | column width overflow Aspose.Cells | C# Aspose.Cells HTML export | unit test tooltip generation | Excel to HTML tooltip overflow
// Common Searches: Aspose.Cells add tooltip for overflow cells | HTML export tooltip only when text is cut off | C# check title attribute in Aspose.Cells HTML output | verify tooltip generation based on column width Aspose | how to test Aspose.Cells HTML tooltip behavior
// Developer Intent: Confirm that the generated HTML contains a title attribute exclusively for cells whose displayed content is clipped by column width.
// Use Cases: Create interactive HTML reports where long values reveal full text on hover while short values stay clean. | Automated regression test to ensure AddTooltipText respects column overflow rules. | Build printable web tables that show tooltips only for truncated entries, improving user experience.
// AI Prompts: Generate a C# unit test that parses the saved HTML and asserts that only overflowed cells have a title attribute. | Extend the sample to detect truncated cells before saving and assign custom tooltip text for each overflowed cell. | Provide a step‑by‑step guide to validate tooltip behavior across multiple rows and columns with varying widths using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace TooltipVerificationDemo
{
    // This C# example creates a workbook with two cells, narrows the first column to force overflow, enables AddTooltipText in HtmlSaveOptions, saves to HTML, and programmatically checks that a <title> attribute is generated for the overflowed cell (A1) while the fitting cell (B1) remains tooltip‑free.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Cell A1: long text that will be truncated
            cells["A1"].PutValue("This is a very long text that will exceed the column width and should show a tooltip.");
            // Cell B1: short text that fits the column
            cells["B1"].PutValue("Short");

            // Set column widths: narrow for A (causing truncation), wide for B (no truncation)
            cells.SetColumnWidth(0, 10); // Column A
            cells.SetColumnWidth(1, 30); // Column B

            // Configure HTML save options to add tooltip text when data is truncated
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.AddTooltipText = true;

            // Save the workbook to HTML
            string htmlPath = "TooltipDemo.html";
            workbook.Save(htmlPath, htmlOptions);

            // Load the generated HTML as plain text
            string htmlContent = File.ReadAllText(htmlPath);

            // Simple verification:
            // Look for a tooltip (title attribute) in the cell representing A1
            bool a1HasTooltip = htmlContent.Contains("<td") && htmlContent.Contains("title=\"This is a very long text");
            // Look for a tooltip in the cell representing B1 (should not exist)
            bool b1HasTooltip = htmlContent.Contains("<td") && htmlContent.Contains(">Short</td") && htmlContent.Contains("title=\"Short\"");

            // Output verification results
            Console.WriteLine("Verification Results:");
            Console.WriteLine($"A1 tooltip present (expected true): {a1HasTooltip}");
            Console.WriteLine($"B1 tooltip present (expected false): {b1HasTooltip}");
        }
    }
}
