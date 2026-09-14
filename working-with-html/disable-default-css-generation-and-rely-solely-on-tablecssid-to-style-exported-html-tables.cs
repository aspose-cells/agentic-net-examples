// Title: Export Excel to HTML with default CSS disabled and a custom TableCssId using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a Workbook as HTML with DisableCss set to true and assigns a custom TableCssId. | Show how to configure HtmlSaveOptions in Aspose.Cells to suppress built‑in CSS and reference an external stylesheet via TableCssId. | Provide a step‑by‑step example of exporting an Excel file to HTML while relying solely on a custom CSS selector.
// Common Searches: how to turn off default CSS when saving Excel as HTML with Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions custom TableCssId for styling exported tables | export Excel workbook to HTML without built‑in styles using C#
// Tags: disable default CSS Aspose.Cells HtmlSaveOptions | custom TableCssId HTML export Aspose.Cells | export Excel to HTML without inline styles | Aspose.Cells DisableCss property usage | apply external stylesheet to Aspose.Cells HTML output

using System;
using Aspose.Cells;

// Loads 'input.xlsx', disables built‑in CSS, sets TableCssId to 'myCustomTable', and saves the workbook as 'output.html' using Aspose.Cells HtmlSaveOptions.
class ExportHtmlWithCustomTableCss
{
    static void Main()
    {
        // Load an existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Disable generation of default CSS styles
            DisableCss = true,

            // Assign a custom CSS ID to the exported HTML table
            TableCssId = "myCustomTable"
        };

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
