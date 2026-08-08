// Title: C# Example – Convert a TSV Workbook to HTML with External CSS using Aspose.Cells
// Description: Demonstrates how to load a tab‑separated values (TSV) file into an Aspose.Cells Workbook, configure HtmlSaveOptions to export worksheet CSS as a separate stylesheet, and save the workbook as an HTML page.
// Keywords: Aspose.Cells TSV to HTML | C# load TSV file | HtmlSaveOptions external CSS | Export worksheet CSS separately | Convert TSV to web page | Aspose.Cells .NET example | TSV to HTML conversion
// Common Searches: Aspose.Cells load TSV in C# | Save workbook as HTML with separate CSS using Aspose.Cells | Export all worksheets to HTML Aspose.Cells .NET | TSV to HTML conversion sample code | How to generate external CSS when saving HTML with Aspose.Cells
// Developer Intent: Load a TSV file into a workbook and export it as an HTML page that references an external CSS stylesheet.
// Use Cases: Create web‑ready reports from TSV data while keeping styling in a maintainable external CSS file. | Publish multi‑worksheet documentation generated from TSV sources with a single shared stylesheet. | Automate batch conversion of TSV datasets to HTML pages for website deployment with centralized style management.
// AI Prompts: Generate C# code that uses Aspose.Cells to read a TSV file and save it as HTML with an external CSS file. | Show how to set HtmlSaveOptions.ExportWorksheetCSSSeparately to true and export all worksheets to HTML. | Explain how to modify the example to embed CSS inline instead of exporting it separately.

using System;
using Aspose.Cells;

// Demonstrates how to load a tab‑separated values (TSV) file into an Aspose.Cells Workbook, configure HtmlSaveOptions to export worksheet CSS as a separate stylesheet, and save the workbook as an HTML page.
class Program
{
    static void Main()
    {
        // Path to the source TSV file
        string tsvFile = "input.tsv";

        // Load the TSV file into a workbook
        LoadOptions loadOpts = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook(tsvFile, loadOpts);

        // Configure HTML save options to export CSS as a separate file
        HtmlSaveOptions htmlOpts = new HtmlSaveOptions();
        htmlOpts.ExportWorksheetCSSSeparately = true;   // external CSS
        htmlOpts.ExportActiveWorksheetOnly = false;    // export all worksheets (optional)

        // Save the workbook as an HTML page
        string htmlFile = "output.html";
        workbook.Save(htmlFile, htmlOpts);

        Console.WriteLine($"HTML page saved to: {htmlFile}");
    }
}
