// Title: Export Excel to a Single‑File HTML with Custom CSS to Override Cell Colors using Aspose.Cells for .NET
// Description: Shows how to load an .xlsx workbook with Aspose.Cells, set HtmlSaveOptions to produce one HTML file, embed custom CSS (via CssStyles) and optionally use CellCssPrefix, then save the workbook so the default cell background and font colors are replaced.
// Keywords: Aspose.Cells | C# | Export Excel to HTML | HtmlSaveOptions | custom CSS | single file HTML | override cell colors | CellCssPrefix | CssStyles | embedded stylesheet
// Common Searches: Aspose.Cells export Excel to HTML with custom stylesheet | C# save workbook as single HTML file Aspose.Cells | how to change cell background color in Aspose.Cells HTML output | HtmlSaveOptions CssStyles example | use CellCssPrefix to style cells in exported HTML
// Developer Intent: The developer needs to convert an Excel workbook into a single HTML document and apply a custom stylesheet that changes the default cell background and text colors.
// Use Cases: Create a web‑ready report that matches corporate branding without external CSS files. | Provide a self‑contained HTML preview of a spreadsheet for email or intranet distribution. | Generate a downloadable HTML version of a workbook with consistent styling across all cells.
// AI Prompts: Generate C# code with Aspose.Cells that saves a workbook as one HTML file and embeds CSS to override cell background and font colors. | Explain the role of HtmlSaveOptions.CellCssPrefix and CssStyles when customizing the HTML output of Aspose.Cells. | Show how to embed a custom stylesheet directly in the exported HTML instead of linking to an external CSS file.

using System;
using Aspose.Cells;

// Shows how to load an .xlsx workbook with Aspose.Cells, set HtmlSaveOptions to produce one HTML file, embed custom CSS (via CssStyles) and optionally use CellCssPrefix, then save the workbook so the default cell background and font colors are replaced.
class ExportExcelToHtmlWithCustomCss
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Save the HTML as a single file so that CssStyles are applied
        saveOptions.SaveAsSingleFile = true;

        // Optional: set a prefix for generated cell CSS classes
        saveOptions.CellCssPrefix = "custom-";

        // Define custom CSS that overrides default cell background and font colors
        saveOptions.CssStyles = @"
            /* Override all table cells */
            td {
                background-color: #e0f7fa !important; /* Light cyan background */
                color: #006064 !important;           /* Dark cyan text */
            }
            /* If cell CSS classes are used, ensure they also get the style */
            .custom-cell {
                background-color: #e0f7fa !important;
                color: #006064 !important;
            }";

        // Save the workbook as HTML with the custom stylesheet
        workbook.Save("output.html", saveOptions);
    }
}
