// Title: Convert Excel to HTML with BestFit layout and _blank links using Aspose.Cells for .NET
// Description: Shows how to load an .xlsx file with Aspose.Cells, enable the BestFit presentation mode, set hyperlinks to open in a new tab (_blank), and save the workbook as an HTML page.
// Keywords: Aspose.Cells | Excel to HTML conversion | HtmlSaveOptions | BestFit layout | PresentationPreference | HtmlLinkTargetType Blank | C# .NET | web‑ready Excel report | auto‑fit columns | _blank hyperlinks
// Common Searches: Aspose.Cells export Excel to HTML with best fit | set hyperlink target _blank when saving Excel as HTML | HtmlSaveOptions PresentationPreference C# example | convert workbook to HTML with auto‑fit columns Aspose | C# code to generate HTML from Excel with blank link target
// Developer Intent: Generate an HTML file from an Excel workbook that automatically adjusts column widths and forces all links to open in a new browser tab.
// Use Cases: Publish Excel‑based dashboards on a website with responsive column sizing. | Create documentation where embedded links must not navigate away from the main page. | Automate batch processing of multiple workbooks into web‑friendly HTML reports.
// AI Prompts: Write C# code with Aspose.Cells to export an Excel workbook to HTML using BestFit layout and set all hyperlinks to open in a new tab. | Explain the impact of HtmlSaveOptions.PresentationPreference and HtmlLinkTargetType.Blank on the resulting HTML. | Provide a PowerShell snippet that calls a .NET assembly to convert Excel files to HTML with blank link targets.

using System;
using Aspose.Cells;

// Shows how to load an .xlsx file with Aspose.Cells, enable the BestFit presentation mode, set hyperlinks to open in a new tab (_blank), and save the workbook as an HTML page.
class ExcelToHtmlConverter
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path where the HTML output will be saved
        string outputPath = "output.html";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(sourcePath);

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Enable presentation preference for a better visual layout (BestFit)
        htmlOptions.PresentationPreference = true;

        // Set hyperlink target to open in a new window or tab (_blank)
        htmlOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine("Workbook has been converted to HTML with PresentationPreference and blank link target.");
    }
}
