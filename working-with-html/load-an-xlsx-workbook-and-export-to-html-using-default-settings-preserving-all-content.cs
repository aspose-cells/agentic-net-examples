// Title: Export XLSX to HTML with Aspose.Cells – default settings preserve all content
// Description: Loads an .xlsx workbook, creates a HtmlSaveOptions object with its out‑of‑the‑box configuration, and saves the workbook as an HTML file while keeping formulas, styles, hidden sheets and other elements intact.
// Keywords: Aspose.Cells export to HTML | XLSX to HTML conversion C# | HtmlSaveOptions default | preserve formulas Excel HTML | convert Excel to web page
// Common Searches: Aspose.Cells convert xlsx to html default options | C# export Excel workbook as HTML preserving styles | How to keep formulas when saving Excel as HTML with Aspose | Save hidden sheets to HTML using Aspose.Cells | Default HtmlSaveOptions behavior Aspose.Cells
// Developer Intent: Generate an HTML representation of an existing XLSX file using Aspose.Cells without customizing any export options.
// Use Cases: Render financial dashboards as web pages while retaining calculation logic. | Provide instant HTML previews of user‑uploaded Excel files in a SaaS portal. | Run a nightly batch job that converts archived .xlsx reports to static HTML archives.
// AI Prompts: Write C# code that loads a .xlsx file and saves it as .html with Aspose.Cells using the default HtmlSaveOptions, ensuring all workbook features are kept. | Explain how Aspose.Cells default HtmlSaveOptions handle formulas, cell styles, and hidden worksheets during Excel‑to‑HTML conversion. | Create a step‑by‑step tutorial for batch converting a directory of .xlsx files to .html in .NET, preserving every workbook element.

using System;
using Aspose.Cells;

// Loads an .xlsx workbook, creates a HtmlSaveOptions object with its out‑of‑the‑box configuration, and saves the workbook as an HTML file while keeping formulas, styles, hidden sheets and other elements intact.
class Program
{
    static void Main()
    {
        // Load the source XLSX workbook. The constructor automatically creates a Workbook instance.
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options with default settings.
        // Default options preserve all content (formulas, styles, hidden sheets, etc.).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Export the workbook to HTML using the default options.
        workbook.Save("output.html", saveOptions);
    }
}
