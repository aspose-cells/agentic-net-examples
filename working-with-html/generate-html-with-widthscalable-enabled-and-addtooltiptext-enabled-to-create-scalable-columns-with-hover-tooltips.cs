// Title: C# – Export Excel to HTML with scalable columns and hover tooltips using Aspose.Cells
// Description: Shows how to build a workbook, shrink a column to force truncation, and save it as HTML with HtmlSaveOptions.WidthScalable and HtmlSaveOptions.AddTooltipText enabled, delivering responsive column widths and mouse‑over tooltips for truncated cells.
// Keywords: Aspose.Cells | HtmlSaveOptions | WidthScalable | AddTooltipText | C# | export Excel to HTML | scalable columns | cell tooltip | truncated content | responsive HTML report
// Common Searches: Aspose.Cells enable WidthScalable when saving HTML | AddTooltipText option for HTML export Aspose.Cells .NET | HTML export with hover tooltips for long cell values | How to make Excel columns scalable in HTML output | Aspose.Cells example showing tooltips for truncated cells
// Developer Intent: Generate an HTML file from a .NET workbook where column widths scale automatically and any truncated cell displays a tooltip on hover.
// Use Cases: Responsive web dashboards that keep layout intact while revealing full text on demand. | Financial or inventory reports where long descriptions are hidden but accessible via tooltips. | Embedding Excel‑derived tables in mobile‑friendly pages without breaking column alignment.
// AI Prompts: Modify the example to apply a custom CSS class to the generated tooltip elements. | Provide a guide for exporting multiple worksheets into one HTML file with WidthScalable and AddTooltipText turned on. | Explain how to change the character limit that triggers tooltip creation in Aspose.Cells HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to build a workbook, shrink a column to force truncation, and save it as HTML with HtmlSaveOptions.WidthScalable and HtmlSaveOptions.AddTooltipText enabled, delivering responsive column widths and mouse‑over tooltips for truncated cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with data
            sheet.Cells["A1"].PutValue("This is a very long text that will not fit in the column width and should show a tooltip.");
            sheet.Cells["B1"].PutValue(12345);

            // Set a narrow column width to force truncation
            sheet.Cells.SetColumnWidth(0, 10); // Column A

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.WidthScalable = true;      // Enable scalable column widths
            htmlOptions.AddTooltipText = true;     // Enable tooltip for truncated content

            // Save the workbook as HTML with the specified options
            string outputPath = "ScalableWithTooltip.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with WidthScalable and AddTooltipText enabled.");
        }
    }
}
