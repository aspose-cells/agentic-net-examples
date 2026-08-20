// Title: Export a Macro‑Free Workbook to HTML with 90% Zoom Using Aspose.Cells for .NET
// Description: Create a macro‑free workbook, set the first worksheet's zoom to 90%, enable WorksheetScalable, and save as HTML while preserving all cell formatting with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# HTML export | worksheet zoom | macro free workbook | HtmlSaveOptions | WorksheetScalable | preserve formatting | export Excel to HTML | PageSetup.Zoom | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set worksheet zoom before HTML export | C# save Excel as HTML with zoom | how to keep formatting when exporting to HTML Aspose.Cells | export macro‑free workbook to HTML .NET | HtmlSaveOptions WorksheetScalable usage | Aspose.Cells HTML export zoom level | C# Aspose.Cells example for HTML output
// Developer Intent: Create a macro‑free Excel workbook, apply a 90% zoom to its first sheet, and export it to HTML while retaining all formatting.
// Use Cases: Generate web‑ready previews of reports where a consistent 90% zoom improves readability. | Provide HTML versions of macro‑free Excel files that exactly match the original layout and styles. | Automate batch conversion of multiple workbooks to HTML with a uniform zoom level and full formatting preservation.
// AI Prompts: Write C# code with Aspose.Cells that creates a new workbook, sets a 90% zoom on the first worksheet, and exports it to HTML preserving all styles. | Explain how HtmlSaveOptions.WorksheetScalable affects the zoom level in the generated HTML and how to disable it if needed. | Show how to modify the example to export only a selected worksheet to HTML while keeping the same zoom and formatting settings.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Create a macro‑free workbook, set the first worksheet's zoom to 90%, enable WorksheetScalable, and save as HTML while preserving all cell formatting with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new, macro‑free workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Apply a 90 % zoom level to the worksheet
            sheet.PageSetup.Zoom = 90;

            // Configure HTML save options to use the worksheet's zoom level
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.WorksheetScalable = true; // Preserve the 90 % zoom in the HTML output

            // Save the workbook as HTML while keeping all formatting
            workbook.Save("output.html", htmlOptions);
        }
    }
}
