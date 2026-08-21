// Title: Convert Excel to HTML with BestFit layout and export comments using Aspose.Cells for .NET
// Description: Loads an .xlsx workbook, enables HtmlSaveOptions.PresentationPreference for BestFit rendering and IsExportComments to include cell comments, then saves the file as an HTML page that preserves visual formatting and annotations.
// Keywords: Aspose.Cells | Excel to HTML conversion | BestFit presentation | PresentationPreference | export cell comments | HtmlSaveOptions | .NET | C# | convert xlsx to html | preserve comments in HTML
// Common Searches: Aspose.Cells convert Excel to HTML with BestFit | Export cell comments when saving workbook as HTML .NET | HtmlSaveOptions PresentationPreference true example | How to keep Excel comments in HTML output using Aspose | BestFit HTML export Aspose.Cells C#
// Developer Intent: Generate an HTML version of an Excel workbook that uses the BestFit visual style and includes every cell comment.
// Use Cases: Create web‑ready reports that look like the original spreadsheet and show comment tooltips. | Provide HTML previews of Excel files for documentation portals while retaining annotations. | Embed spreadsheet data with comments into a web application for interactive help or auditing.
// AI Prompts: Write C# code that converts an Excel workbook to HTML with PresentationPreference set to BestFit and exports all cell comments using Aspose.Cells. | Explain how PresentationPreference and IsExportComments affect the generated HTML and suggest additional HtmlSaveOptions for styling. | Show how to customize the output directory and file naming while preserving comments during HTML conversion with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an .xlsx workbook, enables HtmlSaveOptions.PresentationPreference for BestFit rendering and IsExportComments to include cell comments, then saves the file as an HTML page that preserves visual formatting and annotations.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options:
        // - PresentationPreference = true enables the BestFit (more beautiful) presentation.
        // - IsExportComments = true ensures that cell comments are included in the HTML output.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            PresentationPreference = true,
            IsExportComments = true
        };

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
