// Title: Aspose.Cells C# – Save Workbook as HTML with Hover Tooltips (AddTooltipText)
// Description: Shows how to export an Excel workbook to HTML using Aspose.Cells for .NET, set a narrow column width, and enable the AddTooltipText option so truncated cell values appear as hover tooltips.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | AddTooltipText | HTML export | tooltip | truncated cell | Excel to HTML | hover tooltip | web report
// Common Searches: Aspose.Cells AddTooltipText example C# | export Excel to HTML with tooltips | show full cell value on hover Aspose.Cells | HTML tooltip for truncated cells Aspose | how to enable tooltip text in HtmlSaveOptions
// Developer Intent: Enable hover tooltips that reveal the complete cell content for cells that are truncated in the HTML output.
// Use Cases: Web‑based reports where long descriptions are hidden in narrow columns but accessible via mouse‑over tooltips. | Interactive spreadsheet views in portals that keep column widths compact while still providing full data visibility. | Improved accessibility for exported HTML tables by supplying full cell values through tooltips.
// AI Prompts: Write C# code using Aspose.Cells to save a worksheet as HTML with AddTooltipText set to true so long cell values appear as hover tooltips. | Show how to configure HtmlSaveOptions in Aspose.Cells to automatically add tooltip text for cells truncated in the HTML export. | Provide an example that inserts a long string, narrows a column, and exports to HTML with tooltips enabled using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlTooltipDemo
{
    // Shows how to export an Excel workbook to HTML using Aspose.Cells for .NET, set a narrow column width, and enable the AddTooltipText option so truncated cell values appear as hover tooltips.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into a cell that will exceed the column width
            worksheet.Cells["A1"].PutValue("This is a very long text that will not fit into the cell width and should appear as a tooltip when hovered.");

            // Set a narrow column width to force truncation in the HTML view
            worksheet.Cells.SetColumnWidth(0, 10); // Column A width

            // Configure HTML save options to add tooltip text for truncated cells
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.AddTooltipText = true; // Enable tooltip generation

            // Save the workbook as an HTML file with the specified options
            string outputPath = "HtmlWithTooltip.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved with tooltip enabled: {outputPath}");
        }
    }
}
