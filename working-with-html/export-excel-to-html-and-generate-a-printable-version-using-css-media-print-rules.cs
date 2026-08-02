// Title: Export Excel to HTML with printable CSS (@media print) using Aspose.Cells for .NET
// Description: Loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions to keep grid lines, injects custom CSS that defines @media print rules for hiding non‑print elements and adjusting layout, then saves the workbook as a print‑ready HTML page.
// Keywords: Aspose.Cells HTML export | C# Excel to HTML | printable CSS @media print | export grid lines | custom CssStyles Aspose | .NET spreadsheet to web | HTMLSaveOptions printable
// Common Searches: How to add @media print CSS when saving Excel as HTML with Aspose.Cells | Export Excel to HTML with grid lines and print layout in C# | Aspose.Cells custom CSS for printable HTML output | Generate print‑ready HTML from an Excel workbook .NET
// Developer Intent: Create a web‑friendly HTML version of an Excel workbook that retains grid lines and applies print‑specific styling via CSS.
// Use Cases: Display a spreadsheet on a website while ensuring the printed page matches the on‑screen view. | Produce a downloadable report that hides navigation controls (using a .no‑print class) when printed. | Automate conversion of Excel data into a printable HTML document for offline distribution.
// AI Prompts: Show how to modify the CssStyles string to hide only the header row in the printed HTML. | Generate code that embeds a company logo in the @media print section using Aspose.Cells HtmlSaveOptions. | Explain how to set page size and orientation for the printed HTML via HtmlSaveOptions.PageSetup.

using System;
using Aspose.Cells;

// Loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions to keep grid lines, injects custom CSS that defines @media print rules for hiding non‑print elements and adjusting layout, then saves the workbook as a print‑ready HTML page.
class ExportExcelToHtmlPrintable
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export grid lines so they appear in the printed version
        htmlOptions.ExportGridLines = true;

        // Add custom CSS, including @media print rules for printable output
        htmlOptions.CssStyles = @"
            /* General page styling */
            body { font-family: Arial, sans-serif; margin: 0; padding: 10px; }
            table { border-collapse: collapse; width: 100%; }
            td, th { border: 1px solid #ccc; padding: 5px; }

            /* Print‑specific styling */
            @media print {
                /* Hide elements that should not appear when printing */
                .no-print { display: none !important; }

                /* Ensure the table uses the full printable width */
                table { width: 100% !important; }

                /* Remove page margins for a cleaner print */
                body { margin: 0; }
            }";

        // Save the workbook as an HTML file with the defined options
        workbook.Save("output.html", htmlOptions);
    }
}
