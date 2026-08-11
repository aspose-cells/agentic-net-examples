// Title: C# – Convert XLSX to HTML with Aspose.Cells while Excluding Unused Styles
// Description: Load an XLSX workbook using Aspose.Cells for .NET, enable HtmlSaveOptions.ExcludeUnusedStyles, and save it as HTML to generate a smaller file that contains only the CSS needed for the rendered sheet.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExcludeUnusedStyles | XLSX to HTML | reduce HTML size | minimal CSS export | Excel to HTML conversion | optimize HTML output | Aspose.Cells .NET
// Common Searches: Aspose.Cells exclude unused styles | HtmlSaveOptions ExcludeUnusedStyles C# example | convert Excel workbook to HTML small file | remove unused CSS when exporting XLSX | Aspose.Cells HTML export size reduction | C# export XLSX as HTML minimal CSS
// Developer Intent: Generate an HTML representation of an Excel workbook that omits any CSS rules not applied to the sheet, thereby shrinking the output file.
// Use Cases: Publish lightweight Excel‑derived reports on web pages without excess CSS. | Attach compact HTML previews of large spreadsheets in email communications. | Render workbook snapshots in a web application where bandwidth is limited. | Create archival HTML copies of workbooks while keeping file size minimal.
// AI Prompts: Show how to set HtmlSaveOptions.ExcludeUnusedStyles to true in a C# Aspose.Cells example. | Provide a step‑by‑step C# code snippet that loads an .xlsx file, configures HtmlSaveOptions to drop unused styles, and saves the result as .html. | Explain the impact of ExcludeUnusedStyles on the generated HTML and when it is advisable to use this option. | Give a comparison of HTML file size with and without ExcludeUnusedStyles enabled in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Load an XLSX workbook using Aspose.Cells for .NET, enable HtmlSaveOptions.ExcludeUnusedStyles, and save it as HTML to generate a smaller file that contains only the CSS needed for the rendered sheet.
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string inputPath = "input.xlsx";

            // Path where the HTML file will be saved
            string outputPath = "output.html";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable exclusion of unused styles to reduce the HTML file size
            htmlOptions.ExcludeUnusedStyles = true;

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine("Workbook has been saved as HTML with unused styles excluded.");
        }
    }
}
