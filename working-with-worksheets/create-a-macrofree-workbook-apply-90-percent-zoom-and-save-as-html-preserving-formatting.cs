// Title: C# – Create a Macro‑Free Workbook, Set 90 % Zoom, and Export to HTML with Full Formatting using Aspose.Cells
// Description: Demonstrates how to build a new macro‑free Workbook, apply a 90 % zoom via PageSetup, configure HtmlSaveOptions (WorksheetScalable, ExcludeUnusedStyles, PresentationPreference) and save the sheet as an HTML file that retains all styles and layout.
// Keywords: Aspose.Cells | C# | Workbook | macro‑free workbook | worksheet zoom | 90% zoom | HTML export | HtmlSaveOptions | WorksheetScalable | ExcludeUnusedStyles | PresentationPreference | preserve formatting
// Common Searches: Aspose.Cells set worksheet zoom before HTML export | export workbook to HTML without losing styles | C# example for HtmlSaveOptions WorksheetScalable | how to keep unused styles when converting Excel to HTML | save Excel as HTML with specific zoom level
// Developer Intent: Generate a macro‑free Excel workbook, set a 90 % zoom, and produce an HTML file that keeps every cell style and presentation layout using Aspose.Cells for .NET.
// Use Cases: Create web‑ready reports that match the printed view of a spreadsheet. | Provide an HTML preview of a workbook with a predefined zoom for consistent user experience. | Export macro‑free spreadsheets for embedding in documentation or intranet portals while preserving exact formatting.
// AI Prompts: Write C# code with Aspose.Cells to create a workbook, set a 90 % zoom, and export to HTML preserving all styles. | Show how to configure HtmlSaveOptions (WorksheetScalable, ExcludeUnusedStyles, PresentationPreference) for a full‑format HTML export in Aspose.Cells. | Provide a step‑by‑step example for saving a macro‑free workbook as HTML with scaling based on worksheet zoom.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to build a new macro‑free Workbook, apply a 90 % zoom via PageSetup, configure HtmlSaveOptions (WorksheetScalable, ExcludeUnusedStyles, PresentationPreference) and save the sheet as an HTML file that retains all styles and layout.
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

            // Configure HTML save options to preserve formatting
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Enable scaling based on the worksheet's zoom setting
                WorksheetScalable = true,

                // Keep all styles (do not exclude unused ones) to retain formatting
                ExcludeUnusedStyles = false,

                // Use presentation‑friendly output
                PresentationPreference = true
            };

            // Save the workbook as HTML
            workbook.Save("Workbook90PercentZoom.html", htmlOptions);
        }
    }
}
