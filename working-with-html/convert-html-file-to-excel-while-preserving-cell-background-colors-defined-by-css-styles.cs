// Title: C# – Convert HTML (including CSS background colors) to Excel with Aspose.Cells
// Description: This .NET example demonstrates how to use Aspose.Cells' ConversionUtility.Convert method to read an HTML file that applies CSS for cell background shading and generate an XLSX workbook while keeping the original colors and formatting intact.
// Keywords: Aspose.Cells | C# HTML to Excel conversion | retain CSS cell colors | ConversionUtility Convert | Excel formatting preservation | styled HTML to XLSX | Aspose.Cells .NET library | HTML table to Excel | US C# developers | European .NET community
// Common Searches: Aspose.Cells keep CSS colors during HTML to Excel conversion | C# convert HTML table with background shading to XLSX | How to preserve cell formatting when converting HTML to Excel using Aspose | ConversionUtility retain styles from HTML in Excel | Convert styled HTML invoice to Excel .NET
// Developer Intent: Convert an HTML document to an Excel workbook while retaining CSS‑based cell background colors.
// Use Cases: Generate Excel reports from color‑coded HTML invoices without losing the branding palette. | Migrate web‑based dashboards that use visual cues into Excel spreadsheets for offline analysis. | Provide downloadable .xlsx versions of HTML email newsletters that preserve their original background colors. | Automate extraction of styled HTML tables for financial or statistical modeling in Excel.
// AI Prompts: Show C# code using Aspose.Cells ConversionUtility to convert HTML to XLSX while preserving CSS background colors. | Explain how to configure ConversionUtility to retain cell formatting when converting HTML that references external CSS files. | Give step‑by‑step instructions for handling both inline and linked CSS during HTML‑to‑Excel conversion in .NET. | What options does Aspose.Cells provide to map CSS styles to Excel cell styles?

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// This .NET example demonstrates how to use Aspose.Cells' ConversionUtility.Convert method to read an HTML file that applies CSS for cell background shading and generate an XLSX workbook while keeping the original colors and formatting intact.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Path to the source HTML file (contains CSS background colors)
        string htmlPath = "input.html";

        // Desired path for the resulting Excel workbook
        string excelPath = "output.xlsx";

        // Convert the HTML file to Excel while preserving cell styles (including background colors)
        ConversionUtility.Convert(htmlPath, excelPath);

        Console.WriteLine("HTML successfully converted to Excel. File saved at: " + excelPath);
    }
}
