// Title: Export Excel to HTML with printable CSS @media print using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a print area, and saves it as HTML with HtmlSaveOptions. The export includes grid lines and injects custom CSS that contains @media print rules to hide headers/footers and enforce page breaks, producing a printer‑friendly web page.
// Keywords: Aspose.Cells | C# | .NET | Export Excel to HTML | HTMLSaveOptions | print area | grid lines | CSS @media print | printable HTML | custom CSS | page break
// Common Searches: Aspose.Cells export Excel to HTML printable version | C# save workbook as HTML with print area only | Add @media print CSS when converting Excel to HTML with Aspose | Include grid lines in HTML output from Aspose.Cells | Hide header and footer in HTML export of Excel using Aspose
// Developer Intent: Generate a printer‑friendly HTML file from an Excel workbook that contains only the defined print area, shows grid lines, and applies custom @media print CSS.
// Use Cases: Create a web‑based report that prints cleanly by exporting a specific worksheet range with grid lines and print‑only styling. | Produce a printable financial summary from Excel, hiding UI elements like headers and footers during printing. | Generate multi‑page printable HTML documents from Excel where each table forces a page break using CSS media queries.
// AI Prompts: Show how to modify the CssStyles string to adjust margins, fonts, and colors for the printable HTML version. | Provide code that iterates through all worksheets and saves each as a separate printable HTML file with Aspose.Cells. | Explain how to reference an external CSS file instead of inline CssStyles when exporting Excel to HTML with print media rules.

using System;
using Aspose.Cells;

// Creates a workbook, defines a print area, and saves it as HTML with HtmlSaveOptions. The export includes grid lines and injects custom CSS that contains @media print rules to hide headers/footers and enforce page breaks, producing a printer‑friendly web page.
class ExportExcelToHtmlPrintable
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Report";

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["C1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["C2"].PutValue(0.5);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["C3"].PutValue(0.3);

        // Define the print area (optional, ensures only this range is exported)
        worksheet.PageSetup.PrintArea = "A1:C3";

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export only the defined print area
        htmlOptions.ExportPrintAreaOnly = true;

        // Include grid lines for better visual fidelity
        htmlOptions.ExportGridLines = true;

        // Add custom CSS, including @media print rules for a printable version
        htmlOptions.CssStyles = @"
            /* Styles applied when printing */
            @media print {
                body { margin:0; }
                .header, .footer { display:none; }
                table { page-break-after:always; }
            }

            /* General page styles */
            .header { background:#f2f2f2; padding:10px; }
            .footer { background:#f2f2f2; padding:10px; }
        ";

        // Save the workbook as an HTML file with the specified options
        workbook.Save("Report.html", htmlOptions);
    }
}
