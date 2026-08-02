// Title: Convert TSV to HTML with external CSS using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to load a tab‑separated values (TSV) file into an Aspose.Cells Workbook via LoadOptions (LoadFormat.Tsv) and save it as an HTML page. HtmlSaveOptions are configured with ExportWorksheetCSSSeparately so the stylesheet is written to a separate .css file while all worksheets are included.
// Keywords: Aspose.Cells | C# | TSV to HTML | ExportWorksheetCSSSeparately | external CSS | HtmlSaveOptions | LoadFormat.Tsv | convert TSV | save workbook as HTML | separate stylesheet
// Common Searches: Aspose.Cells load TSV file C# | Save workbook as HTML with external CSS Aspose | Export worksheet CSS separately Aspose.Cells | Convert tab separated values to HTML C# | How to generate HTML and CSS from TSV using Aspose.Cells
// Developer Intent: Load a TSV workbook and export it to an HTML file that references a separate CSS stylesheet.
// Use Cases: Create web‑ready reports from TSV data with styling kept in an external CSS file for easy maintenance. | Publish multiple worksheets from a TSV source as a single HTML page with a shared stylesheet to leverage browser caching. | Automate batch conversion of TSV files to HTML for content management systems while preserving formatting via external CSS.
// AI Prompts: Generate C# code that uses Aspose.Cells to read a TSV file and save it as HTML with the stylesheet saved separately. | Describe how the ExportWorksheetCSSSeparately option changes the output files when saving a workbook as HTML. | Provide a step‑by‑step tutorial for converting a TSV workbook to HTML and linking the generated CSS file in the HTML head.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTsvToHtml
{
    // This C# example shows how to load a tab‑separated values (TSV) file into an Aspose.Cells Workbook via LoadOptions (LoadFormat.Tsv) and save it as an HTML page. HtmlSaveOptions are configured with ExportWorksheetCSSSeparately so the stylesheet is written to a separate .css file while all worksheets are included.
    class Program
    {
        static void Main()
        {
            // Path to the source TSV file
            string tsvPath = Path.Combine(Environment.CurrentDirectory, "input.tsv");

            // Load the TSV file into a workbook using LoadOptions with Tsv format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Configure HTML save options to export worksheet CSS separately
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportWorksheetCSSSeparately = true;   // CSS will be saved as a separate .css file
            htmlOptions.ExportActiveWorksheetOnly = false;    // Export all worksheets (default)

            // Define the output HTML file path
            string htmlPath = Path.Combine(Environment.CurrentDirectory, "output.html");

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine("TSV file has been converted to HTML.");
            Console.WriteLine("HTML file: " + htmlPath);
            Console.WriteLine("CSS file is generated alongside the HTML file.");
        }
    }
}
