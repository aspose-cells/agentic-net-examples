// Title: Export Excel to HTML with BestFit layout and CSS custom properties for image deduplication (Aspose.Cells C#)
// Description: This C# example loads an .xlsx file, sets HtmlSaveOptions.PresentationPreference to BestFit for a compact, responsive layout, enables EnableCssCustomProperties to consolidate repeated base64 images, and saves the workbook as an optimized HTML page.
// Keywords: Aspose.Cells | Excel to HTML | PresentationPreference BestFit | EnableCssCustomProperties | image deduplication | C# export | HTML size reduction | base64 image reuse
// Common Searches: Aspose.Cells export Excel to HTML BestFit | How to deduplicate images in HTML output using Aspose.Cells | Enable CSS custom properties in Aspose.Cells C# | Reduce HTML size when converting Excel with Aspose.Cells | PresentationPreference option Aspose.Cells example
// Developer Intent: Create an HTML version of an Excel workbook that preserves layout while eliminating duplicate image data.
// Use Cases: Generate lightweight web reports from Excel files for intranet portals. | Automate batch conversion of multiple workbooks to bandwidth‑efficient HTML pages. | Embed Excel‑derived tables in SaaS dashboards where responsive design and fast load times are critical.
// AI Prompts: Provide a C# example that configures HtmlSaveOptions.PresentationPreference = true and EnableCssCustomProperties = true for Aspose.Cells. | Explain the impact of PresentationPreference and EnableCssCustomProperties on the generated HTML and how to adjust them for different devices. | Show a script to process a directory of .xlsx files into HTML using Aspose.Cells while removing duplicate base64 images.

using System;
using Aspose.Cells;

// This C# example loads an .xlsx file, sets HtmlSaveOptions.PresentationPreference to BestFit for a compact, responsive layout, enables EnableCssCustomProperties to consolidate repeated base64 images, and saves the workbook as an optimized HTML page.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Enable presentation preference (BestFit) for a more beautiful layout
        htmlOptions.PresentationPreference = true;

        // Enable CSS custom properties to deduplicate repeated base64 images
        htmlOptions.EnableCssCustomProperties = true;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
