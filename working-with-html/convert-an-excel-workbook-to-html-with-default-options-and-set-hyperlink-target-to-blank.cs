// Title: Convert Excel to HTML with _blank hyperlinks using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file, adds a sample hyperlink, configures HtmlSaveOptions to set LinkTargetType = Blank, and saves the workbook as an HTML page where all links open in a new browser tab.
// Keywords: Aspose.Cells | C# Excel to HTML | HtmlSaveOptions | LinkTargetType Blank | hyperlink target _blank | convert .xlsx to .html | export workbook as HTML | Aspose.Cells hyperlink | HTML export options | new tab links
// Common Searches: Aspose.Cells export Excel to HTML _blank | How to set hyperlink target when saving Excel as HTML in C# | HtmlSaveOptions LinkTargetType example | Convert .xlsx to HTML with Aspose.Cells .NET | C# code to open Excel hyperlinks in new tab after HTML conversion
// Developer Intent: Generate an HTML version of an Excel workbook where every hyperlink uses the _blank target so it opens in a new browser tab.
// Use Cases: Web dashboards that embed spreadsheet data with external links opening in separate tabs. | Email or intranet reports that need an HTML preview of Excel files with safe link behavior. | Automated batch conversion of multiple .xlsx files to .html while enforcing a uniform _blank link target. | Documentation generation from Excel spreadsheets with consistent hyperlink handling.
// AI Prompts: Write C# code using Aspose.Cells to convert an Excel workbook to HTML and set all hyperlink targets to _blank. | Explain the purpose of HtmlSaveOptions.LinkTargetType and how it affects hyperlink rendering in the generated HTML. | Create a script that processes every .xlsx file in a folder, converting each to HTML with Aspose.Cells and applying the _blank link target. | Show how to customize the output folder and CSS while preserving _blank hyperlinks during Excel‑to‑HTML conversion.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Loads an .xlsx file, adds a sample hyperlink, configures HtmlSaveOptions to set LinkTargetType = Blank, and saves the workbook as an HTML page where all links open in a new browser tab.
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your actual file path)
            string excelPath = "input.xlsx";
            Workbook workbook = new Workbook(excelPath);

            // Example: add a hyperlink to demonstrate the target attribute
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Visit Aspose");
            sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

            // Create HTML save options with default settings
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Set hyperlink target to open in a new window/tab (_blank)
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

            // Save the workbook as HTML using the configured options
            string htmlPath = "output.html";
            workbook.Save(htmlPath, saveOptions);

            Console.WriteLine($"Workbook successfully converted to HTML: {htmlPath}");
        }
    }
}
