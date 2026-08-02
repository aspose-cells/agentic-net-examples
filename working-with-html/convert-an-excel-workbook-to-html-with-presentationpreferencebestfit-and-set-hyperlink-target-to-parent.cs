// Title: Convert Excel to HTML with BestFit layout and parent hyperlink target – Aspose.Cells for .NET
// Description: Shows how to load an .xlsx workbook with Aspose.Cells, enable PresentationPreference (BestFit) for optimal column widths, set HtmlLinkTargetType to Parent so links open in the same window, and save the workbook as an HTML file.
// Keywords: Aspose.Cells HTML export | PresentationPreference BestFit | HtmlLinkTargetType Parent | C# convert Excel to HTML | Excel to web‑ready HTML | preserve column widths Aspose | hyperlink target parent | Aspose.Cells .NET example
// Common Searches: Aspose.Cells export Excel to HTML BestFit | Set hyperlink target to parent in Aspose.Cells HTML | HtmlSaveOptions PresentationPreference C# | How to keep column widths when converting Excel to HTML | Aspose.Cells HTML link target parent frame | C# code to convert .xlsx to HTML with best fit
// Developer Intent: Export an Excel workbook to HTML with column widths automatically adjusted (BestFit) and hyperlinks that open in the parent frame.
// Use Cases: Create web‑ready reports from Excel while maintaining original column sizing. | Embed HTML snippets in dashboards where links must navigate the surrounding page. | Batch‑process multiple spreadsheets into consistent HTML files for intranet publishing.
// AI Prompts: Generate C# code using Aspose.Cells to convert an Excel file to HTML with PresentationPreference set to BestFit and LinkTargetType set to Parent. | Explain the impact of HtmlSaveOptions.PresentationPreference and HtmlLinkTargetType on the generated HTML output. | Write a PowerShell script that invokes a .NET assembly to batch‑convert all .xlsx files in a directory to HTML using the same BestFit and parent link settings.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to load an .xlsx workbook with Aspose.Cells, enable PresentationPreference (BestFit) for optimal column widths, set HtmlLinkTargetType to Parent so links open in the same window, and save the workbook as an HTML file.
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Enable presentation preference for a better visual layout (BestFit)
            saveOptions.PresentationPreference = true;

            // Ensure hyperlinks open in the parent frame (default is Parent, set explicitly)
            saveOptions.LinkTargetType = HtmlLinkTargetType.Parent;

            // Save the workbook as an HTML file with the specified options
            workbook.Save("output.html", saveOptions);
        }
    }
}
