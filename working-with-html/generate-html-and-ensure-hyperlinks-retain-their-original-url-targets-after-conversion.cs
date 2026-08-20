// Title: Export Excel to HTML with original hyperlink targets using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, adds an external hyperlink with optional display text, sets HtmlSaveOptions.LinkTargetType to Self so the link keeps its original URL, and saves the workbook as an HTML file.
// Keywords: Aspose.Cells | HtmlSaveOptions | LinkTargetType | Self | preserve hyperlink | C# | export Excel to HTML | hyperlink target | .NET | workbook HTML conversion
// Common Searches: Aspose.Cells keep hyperlink target when saving as HTML | HtmlSaveOptions LinkTargetType Self C# example | Export Excel workbook to HTML with original URLs | Set hyperlink display text in Aspose.Cells | C# save workbook as HTML preserving external links
// Developer Intent: Generate an HTML file from an Excel workbook while ensuring hyperlinks retain their original URLs and open in the same window.
// Use Cases: Add external links to cells and export the sheet to HTML for web reporting. | Show custom link text in the HTML output without altering the underlying URL. | Control link behavior (same‑window vs. new tab) by configuring HtmlSaveOptions. | Create static HTML dashboards from Excel data that preserve navigation paths.
// AI Prompts: Provide a C# example that exports an Aspose.Cells workbook to HTML with hyperlinks that keep their original targets. | Explain the effect of HtmlSaveOptions.LinkTargetType = Self on hyperlink behavior in the generated HTML. | Show how to add a hyperlink with custom display text to a cell and retain it during HTML conversion using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new workbook, adds an external hyperlink with optional display text, sets HtmlSaveOptions.LinkTargetType to Self so the link keeps its original URL, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell A1 pointing to an external URL
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

        // Set the display text for the hyperlink (optional)
        worksheet.Cells["A1"].PutValue("Visit Example");

        // Configure HTML save options to keep the original hyperlink target (open in the same window)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Self; // retains original URL target

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", saveOptions);
    }
}
