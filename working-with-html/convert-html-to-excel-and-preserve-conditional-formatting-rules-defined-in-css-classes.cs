// Title: C# – Convert HTML to Excel while retaining CSS conditional formatting with Aspose.Cells
// Description: Learn how to load an HTML file using Aspose.Cells LoadOptions (Html) and convert it to an XLSX workbook via ConversionUtility. The process automatically maps CSS class styles, including conditional‑formatting rules, so the resulting spreadsheet mirrors the original HTML appearance.
// Keywords: Aspose.Cells HTML to Excel conversion | C# preserve CSS conditional formatting | ConversionUtility Excel export | LoadOptions Html Aspose | map CSS classes to Excel styles
// Common Searches: Aspose.Cells keep CSS formatting when converting HTML to XLSX | C# convert HTML file to Excel preserving conditional rules | How to map HTML CSS classes to Excel styles with Aspose | HTML to Excel conversion with conditional formatting in .NET
// Developer Intent: Transform an HTML document into an Excel workbook without losing any CSS‑driven conditional formatting.
// Use Cases: Export web‑based financial dashboards that use CSS highlights into Excel for offline analysis. | Migrate HTML email reports containing data‑driven color rules to spreadsheet format. | Automate generation of Excel templates from HTML marketing assets while preserving visual cues.
// AI Prompts: Show how to customize OoxmlSaveOptions (e.g., set compression level) while still keeping CSS conditional formatting. | Provide code that reads the HTML from a MemoryStream instead of a file and retains formatting during conversion. | Explain how to programmatically inspect the workbook to confirm that conditional‑formatting rules were transferred.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Learn how to load an HTML file using Aspose.Cells LoadOptions (Html) and convert it to an XLSX workbook via ConversionUtility. The process automatically maps CSS class styles, including conditional‑formatting rules, so the resulting spreadsheet mirrors the original HTML appearance.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Path to the source HTML file that contains CSS classes for conditional formatting
        string sourceHtml = "input.html";

        // Path where the resulting Excel file will be saved
        string destinationXlsx = "output.xlsx";

        // LoadOptions tells Aspose.Cells to interpret the source file as HTML
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

        // When loading HTML, Aspose.Cells automatically maps CSS class based styles
        // (including conditional formatting) to the corresponding Excel style objects.
        // No additional flags are required; just ensure that the HTML is well‑formed.

        // Convert the HTML file to an Excel workbook using the ConversionUtility.
        // OoxmlSaveOptions can be used to control Excel saving, but default options are sufficient.
        SaveOptions saveOptions = new OoxmlSaveOptions();

        // Perform the conversion
        ConversionUtility.Convert(sourceHtml, loadOptions, destinationXlsx, saveOptions);

        Console.WriteLine("HTML has been successfully converted to Excel with CSS‑based conditional formatting preserved.");
    }
}
