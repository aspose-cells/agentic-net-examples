// Title: C# – Convert Excel to HTML with PresentationPreference (BestFit) and Parent hyperlink target using Aspose.Cells
// Description: Loads an .xlsx workbook, enables HtmlSaveOptions.PresentationPreference for a best‑fit HTML layout, sets HtmlLinkTargetType to Parent so links open in the parent frame, and saves the file as HTML.
// Keywords: Aspose.Cells | C# Excel to HTML | HtmlSaveOptions PresentationPreference | BestFit HTML export | HtmlLinkTargetType Parent | preserve column widths | web‑ready HTML report | Excel workbook conversion
// Common Searches: Aspose.Cells export Excel to HTML best fit | set hyperlink target parent Aspose.Cells HTML | C# HtmlSaveOptions PresentationPreference example | convert .xlsx to HTML with Aspose.Cells | HTML output preserving layout Aspose.Cells
// Developer Intent: Generate an HTML file from an Excel workbook that keeps the original column widths and layout while making all hyperlinks open in the parent frame.
// Use Cases: Create web‑ready reports from Excel files that retain the exact spreadsheet layout. | Embed generated HTML into portals or dashboards where link clicks must stay within the surrounding page. | Automate batch conversion of multiple .xlsx files to consistently styled HTML with parent‑frame link behavior.
// AI Prompts: Write C# code using Aspose.Cells to export an Excel workbook to HTML with PresentationPreference enabled and link target set to parent. | Explain how HtmlSaveOptions.PresentationPreference affects the HTML output and how to configure HtmlLinkTargetType in Aspose.Cells. | Provide a step‑by‑step guide to batch‑process a folder of .xlsx files into best‑fit HTML pages with parent hyperlink targets.

using System;
using Aspose.Cells;

// Loads an .xlsx workbook, enables HtmlSaveOptions.PresentationPreference for a best‑fit HTML layout, sets HtmlLinkTargetType to Parent so links open in the parent frame, and saves the file as HTML.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Enable presentation preference for a more beautiful HTML output
        saveOptions.PresentationPreference = true;

        // Set hyperlink target to open in the parent frame
        saveOptions.LinkTargetType = HtmlLinkTargetType.Parent;

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
