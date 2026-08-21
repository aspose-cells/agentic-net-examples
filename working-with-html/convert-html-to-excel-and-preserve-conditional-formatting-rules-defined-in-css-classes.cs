// Title: C# – Convert HTML with CSS Conditional Formatting to Excel using Aspose.Cells
// Description: A .NET example that loads an HTML file containing CSS‑based conditional formatting via LoadOptions (Html) and converts it to an XLSX workbook with ConversionUtility, automatically preserving the formatting rules.
// Keywords: Aspose.Cells | C# | HTML to Excel conversion | CSS conditional formatting | LoadOptions Html | ConversionUtility | preserve styles | Excel export .NET | GitHub example | Aspose.Cells API
// Common Searches: Aspose.Cells preserve CSS conditional formatting when converting HTML to Excel | C# convert HTML table with conditional formatting to XLSX | Load HTML with CSS styles into workbook Aspose.Cells | ConversionUtility HTML to Excel example | How to keep conditional formatting from HTML in Excel using Aspose
// Developer Intent: Transform an HTML document that uses CSS conditional formatting into an Excel workbook while retaining the original formatting rules.
// Use Cases: Export web‑based financial dashboards that rely on CSS conditional formatting into Excel for further analysis. | Migrate styled HTML reports or email templates into Excel without losing visual cues such as color‑coded thresholds. | Automate bulk conversion of HTML tables with conditional styling into XLSX files for downstream data processing.
// AI Prompts: Show how to map specific CSS classes to Excel conditional formatting after using ConversionUtility. | Provide code to validate that conditional formatting from the source HTML was retained in the generated workbook. | Explain how to customize LoadOptions to handle external CSS files during HTML‑to‑Excel conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A .NET example that loads an HTML file containing CSS‑based conditional formatting via LoadOptions (Html) and converts it to an XLSX workbook with ConversionUtility, automatically preserving the formatting rules.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Path to the source HTML file that contains CSS‑based conditional formatting
        string htmlFile = "input.html";

        // Desired path for the resulting Excel workbook
        string excelFile = "output.xlsx";

        // LoadOptions specify that the source file is HTML.
        // This enables Aspose.Cells to parse the HTML and its CSS styles.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

        // OoxmlSaveOptions are used for saving the workbook in XLSX format.
        // No special settings are required for preserving conditional formatting,
        // as the parsing of CSS classes is handled during the load phase.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();

        // Perform the conversion from HTML to Excel.
        // The ConversionUtility respects the provided load and save options.
        ConversionUtility.Convert(htmlFile, loadOptions, excelFile, saveOptions);

        Console.WriteLine("Conversion completed. Excel file saved to: " + excelFile);
    }
}
