// Title: C# – Convert HTML to Excel with page breaks using Aspose.Cells
// Description: Shows how to use Aspose.Cells.Utility.ConversionUtility.Convert in C# to turn an HTML file into an XLSX workbook while automatically mapping HTML page‑break elements (<hr> tags or CSS page‑break properties) to worksheet page breaks.
// Keywords: Aspose.Cells | HTML to Excel conversion | C# | ConversionUtility | preserve page breaks | worksheet page breaks | HTML <hr> to Excel | CSS page-break | batch HTML conversion | pagination preservation
// Common Searches: Aspose.Cells convert HTML to Excel with page breaks | C# keep <hr> as Excel page break | preserve CSS page-break when converting HTML to XLSX | ConversionUtility page break mapping | HTML to Excel pagination using .NET
// Developer Intent: Convert an HTML document to an Excel workbook while retaining the original HTML page‑break markers as worksheet page breaks.
// Use Cases: Create printable Excel reports from web pages that use <hr> or CSS page‑breaks, keeping the same pagination. | Automate conversion of HTML invoices that rely on page‑break‑after rules to Excel for downstream processing. | Batch‑process multiple HTML files into Excel workbooks, preserving pagination for archiving or analysis.
// AI Prompts: Generate C# code that adds a custom page break after a specific HTML element during conversion with Aspose.Cells. | Provide a snippet to read the resulting workbook and list all worksheet page break positions after HTML conversion. | Explain how to configure ConversionUtility to ignore certain HTML tags while still preserving page breaks.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Shows how to use Aspose.Cells.Utility.ConversionUtility.Convert in C# to turn an HTML file into an XLSX workbook while automatically mapping HTML page‑break elements (<hr> tags or CSS page‑break properties) to worksheet page breaks.
class HtmlToExcelWithPageBreaks
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path where the resulting Excel file will be saved
        string excelPath = "output.xlsx";

        // Convert the HTML document to an Excel workbook.
        // Aspose.Cells automatically maps HTML page‑break elements (e.g., <hr> or CSS page‑break) 
        // to worksheet page breaks during the conversion.
        ConversionUtility.Convert(htmlPath, excelPath);

        Console.WriteLine("Conversion completed. Excel file saved to: " + excelPath);
    }
}
