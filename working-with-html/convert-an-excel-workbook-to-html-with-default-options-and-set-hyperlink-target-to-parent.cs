// Title: Export Excel to HTML with Parent‑Frame Hyperlinks using Aspose.Cells for .NET (C#)
// Description: Shows how to load an .xlsx workbook, configure HtmlSaveOptions to set LinkTargetType to Parent (the default), and save the workbook as HTML so every hyperlink opens in the parent window or frame.
// Keywords: Aspose.Cells HTML export | C# convert Excel to HTML | HtmlSaveOptions LinkTargetType | hyperlink target parent | default HTML save options | Aspose.Cells .NET | Excel to web report | parent frame links
// Common Searches: Aspose.Cells export Excel to HTML C# | set hyperlink target parent Aspose.Cells | HtmlSaveOptions LinkTargetType Parent example | convert workbook to HTML with default options | how to make Excel links open in parent frame
// Developer Intent: Generate an HTML representation of an Excel workbook where all links open in the surrounding page rather than a new tab or window.
// Use Cases: Publish Excel‑based dashboards on a website and keep navigation within the main page. | Embed HTML spreadsheets inside an iframe while allowing links to affect the parent document. | Automate batch conversion of multiple workbooks to HTML with consistent parent‑frame hyperlink behavior.
// AI Prompts: Provide C# code that converts an Excel file to HTML with Aspose.Cells and forces hyperlinks to open in the parent frame. | Explain the effect of HtmlLinkTargetType.Parent on generated HTML links in Aspose.Cells. | Show a minimal example of using HtmlSaveOptions to save a workbook as HTML while customizing the link target.

using System;
using Aspose.Cells;

// Shows how to load an .xlsx workbook, configure HtmlSaveOptions to set LinkTargetType to Parent (the default), and save the workbook as HTML so every hyperlink opens in the parent window or frame.
class Program
{
    static void Main()
    {
        // Load an existing Excel workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and set hyperlink target to the parent frame
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Parent; // default, set explicitly

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", saveOptions);
    }
}
