// Title: Export Excel to HTML with Aspose.Cells – set hyperlinks to open in new tabs (C#)
// Description: Load an .xlsx file using Aspose.Cells, apply HtmlSaveOptions.LinkTargetType = HtmlLinkTargetType.Blank, and save as HTML so every link renders with target="_blank".
// Keywords: Aspose.Cells | Excel to HTML conversion | HtmlSaveOptions | LinkTargetType | HtmlLinkTargetType.Blank | C# workbook export | hyperlink target blank | web‑ready spreadsheet | save workbook as HTML | open links in new tab
// Common Searches: Aspose.Cells export Excel as HTML C# | HtmlSaveOptions LinkTargetType example | make Excel hyperlinks open in new window when saved as HTML | target='_blank' with Aspose.Cells HTML output | C# convert .xlsx to web page with clickable links
// Developer Intent: Generate an HTML file from an Excel workbook and configure all embedded links to launch in a separate browser tab.
// Use Cases: Publish a spreadsheet on a website while keeping external references from navigating away from the page. | Create documentation that includes Excel data with links that open in their own tabs for smoother user experience. | Integrate Excel‑derived content into a web portal, preserving link behavior that opens in new windows.
// AI Prompts: Show a C# example that converts an Excel file to HTML with Aspose.Cells and forces every hyperlink to use target='_blank'. | Explain how setting HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank changes the generated HTML. | Provide step‑by‑step code to load a workbook, configure link targets, and save as HTML using Aspose.Cells.

using System;
using Aspose.Cells;

// Load an .xlsx file using Aspose.Cells, apply HtmlSaveOptions.LinkTargetType = HtmlLinkTargetType.Blank, and save as HTML so every link renders with target="_blank".
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(sourcePath);

        // Initialize HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Ensure that all hyperlinks in the generated HTML open in a new tab/window
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Define the output HTML file path
        string outputPath = "output.html";

        // Save the workbook as an HTML file using the configured options
        workbook.Save(outputPath, saveOptions);
    }
}
