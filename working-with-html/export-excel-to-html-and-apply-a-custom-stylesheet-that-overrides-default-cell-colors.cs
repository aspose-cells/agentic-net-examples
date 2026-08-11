// Title: Export Excel to HTML with a custom CSS stylesheet that overrides cell colors using Aspose.Cells for .NET
// Description: Loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions to embed a custom CSS string that changes background and text colors for all <td> elements (and optional classes), ensures CSS is applied by keeping DisableCss false, and saves the workbook as an HTML file with the new styling.
// Keywords: Aspose.Cells | C# | Excel to HTML | custom CSS stylesheet | override cell colors | HtmlSaveOptions | CssStyles property | DisableCss false | web report styling | brand colors in HTML export
// Common Searches: Aspose.Cells export Excel as HTML with custom CSS | change cell background color in HTML output using Aspose.Cells | apply external stylesheet when saving workbook to HTML .NET | override default cell styles in Aspose.Cells HTML export | how to use HtmlSaveOptions.CssStyles in C#
// Developer Intent: Create an HTML version of an Excel workbook and apply a custom CSS stylesheet that replaces the default cell coloring.
// Use Cases: Publish Excel‑based dashboards on a website with a unified brand color scheme. | Generate HTML reports where all cells share a specific background and text color without altering the source workbook. | Add CSS classes to selected cells for conditional formatting that appears only in the exported HTML page.
// AI Prompts: Show how to load a CSS file from disk and assign its contents to HtmlSaveOptions.CssStyles before exporting to HTML. | Provide C# code that adds different CSS classes to rows based on their numeric values, then saves the workbook as HTML. | Explain how to disable inline styles completely and reference an external CSS file when saving an Excel workbook to HTML with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions to embed a custom CSS string that changes background and text colors for all <td> elements (and optional classes), ensures CSS is applied by keeping DisableCss false, and saves the workbook as an HTML file with the new styling.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path for the generated HTML file
        string htmlPath = "output.html";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(excelPath);

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Define custom CSS that overrides default cell colors
        // This CSS will be applied to all table cells (td) and can be extended as needed
        saveOptions.CssStyles = @"
            td {
                background-color: #ffebcd !important;   /* Light orange background for all cells */
                color: #2f4f4f !important;              /* Dark slate gray text color */
            }
            .custom-cell {
                background-color: #add8e6 !important;   /* Light blue for cells with this class */
                color: #000080 !important;              /* Navy text color */
            }";

        // Ensure that CSS styles are used (not only inline styles)
        saveOptions.DisableCss = false;

        // Save the workbook as an HTML file with the custom stylesheet
        workbook.Save(htmlPath, saveOptions);

        Console.WriteLine($"HTML file saved to: {htmlPath}");
    }
}
