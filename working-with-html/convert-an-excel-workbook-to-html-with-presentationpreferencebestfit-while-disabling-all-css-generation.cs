// Title: C# – Export Excel to HTML with BestFit layout and inline styles only using Aspose.Cells
// Description: Demonstrates loading an Excel workbook, configuring HtmlSaveOptions to enable PresentationPreference (BestFit) and disable all CSS generation, then saving the file as HTML with only inline styling.
// Keywords: Aspose.Cells | C# HTML export | PresentationPreference BestFit | DisableCss | Excel to HTML | inline styles only | HtmlSaveOptions | .NET spreadsheet conversion
// Common Searches: Aspose.Cells export Excel to HTML BestFit | disable CSS when saving workbook as HTML Aspose.Cells | C# convert Excel to HTML with inline styles | HtmlSaveOptions PresentationPreference true | Aspose.Cells HTML output without external CSS
// Developer Intent: Create an HTML representation of an Excel workbook that uses the BestFit presentation mode and contains no external CSS files, only inline styling, via Aspose.Cells for .NET.
// Use Cases: Generate web‑ready spreadsheet views that preserve column widths while keeping all styling inside the HTML markup. | Produce email‑compatible HTML reports where external CSS files are blocked. | Archive workbook snapshots for documentation or version control without managing separate style sheets.
// AI Prompts: Show how to set HtmlSaveOptions.PresentationPreference to BestFit and disable CSS generation in Aspose.Cells C#. | Provide a complete C# example that loads an .xlsx file and saves it as HTML with only inline styles using Aspose.Cells. | Explain the effect of HtmlSaveOptions.DisableCss and PresentationPreference on the structure and size of the generated HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates loading an Excel workbook, configuring HtmlSaveOptions to enable PresentationPreference (BestFit) and disable all CSS generation, then saving the file as HTML with only inline styling.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            // If you don't have an input file, you can create a new workbook instead.
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable presentation preference for a more beautiful layout (BestFit)
            htmlOptions.PresentationPreference = true;

            // Disable all CSS generation – only inline styles will be used
            htmlOptions.DisableCss = true;

            // Save the workbook as HTML
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook has been converted to HTML with PresentationPreference and CSS disabled.");
        }
    }
}
