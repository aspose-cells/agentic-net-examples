// Title: Export Excel to HTML without tooltips using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export workbook to HTML with AddTooltipText set to false | how to improve HTML export speed in Aspose.Cells by disabling tooltips | Aspose.Cells HtmlSaveOptions disable tooltip text example | generate HTML from Excel in .NET without tooltip pop‑ups
// Tags: Aspose.Cells HtmlSaveOptions disable tooltips | C# Excel to HTML conversion performance | AddTooltipText false Aspose.Cells | HTML export speed optimization Aspose.Cells | save workbook as HTML without tooltip text

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, populating cells, configuring HtmlSaveOptions with AddTooltipText = false to disable tooltip generation, and saving the workbook as an HTML file for faster rendering.
class HtmlExportWithNoTooltips
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // creates an empty workbook

        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["A3"].PutValue(456);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Disable tooltip generation to improve rendering speed
            AddTooltipText = false
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
