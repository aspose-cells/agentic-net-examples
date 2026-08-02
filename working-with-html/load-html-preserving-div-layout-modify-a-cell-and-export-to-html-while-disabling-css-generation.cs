// Title: Load HTML with DIV layout, edit a cell, and save as CSS‑free HTML using Aspose.Cells for .NET
// Description: Demonstrates how to load an HTML file into a Workbook with HtmlLoadOptions.SupportDivTag to keep the original <div> structure, modify a cell value, and export the workbook back to HTML with HtmlSaveOptions.DisableCss so only inline styles are generated.
// Keywords: Aspose.Cells | C# | .NET | HtmlLoadOptions | SupportDivTag | HtmlSaveOptions | DisableCss | load HTML workbook | edit cell value | export HTML without CSS | inline styles
// Common Searches: Aspose.Cells load HTML preserving div layout | How to disable CSS when saving HTML with Aspose.Cells | Modify cell after loading HTML workbook in C# | Save Aspose.Cells workbook to HTML with inline styles only | SupportDivTag example Aspose.Cells .NET
// Developer Intent: Load an HTML file while keeping its DIV layout, change a cell, and save the result as HTML without external CSS.
// Use Cases: Refresh data in an existing HTML‑based report without altering its page layout. | Generate email‑ready HTML output with inline styling after programmatic cell updates. | Batch‑process legacy HTML tables, update specific cells, and produce CSS‑free HTML for web deployment.
// AI Prompts: Write C# code that loads an HTML file with SupportDivTag enabled, updates cell B2, and saves the workbook to HTML with DisableCss set to true using Aspose.Cells. | Explain how HtmlLoadOptions.SupportDivTag and HtmlSaveOptions.DisableCss affect the DOM and styling of the exported HTML. | Create a script that iterates over multiple HTML files, changes a designated cell in each workbook, and exports each to CSS‑free HTML with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to load an HTML file into a Workbook with HtmlLoadOptions.SupportDivTag to keep the original <div> structure, modify a cell value, and export the workbook back to HTML with HtmlSaveOptions.DisableCss so only inline styles are generated.
class Program
{
    static void Main()
    {
        // Load the HTML file and enable support for <div> tags to preserve layout
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.SupportDivTag = true;
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Modify a cell (example: set A1 to a new value)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Modified Value");

        // Save the workbook as HTML with CSS disabled (inline styles only)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DisableCss = true;
        workbook.Save("output.html", saveOptions);
    }
}
