// Title: C# – Convert HTML to PDF with Precise Table Border Widths via Aspose.Cells
// Description: This example loads an HTML file into an Aspose.Cells Workbook, disables border collapsing (HtmlSaveOptions.IsBorderCollapsed = false) and enables ExportSimilarBorderStyle to retain each border’s pixel thickness. After optionally saving an intermediate HTML file, the workbook is rendered to PDF, producing a document where table borders match the original HTML dimensions.
// Keywords: Aspose.Cells | HTML to PDF | C# .NET | table border width | IsBorderCollapsed | ExportSimilarBorderStyle | preserve border thickness | PDF rendering | Aspose.Cells HtmlSaveOptions | convert HTML workbook to PDF
// Common Searches: Aspose.Cells keep HTML table borders when converting to PDF | HtmlSaveOptions IsBorderCollapsed false example | ExportSimilarBorderStyle C# Aspose.Cells | HTML to PDF conversion preserving border thickness | C# convert HTML file to PDF Aspose.Cells
// Developer Intent: Create a PDF from an HTML source while preserving the original pixel‑level border styling of tables.
// Use Cases: Generating printable invoices from web pages where border lines must stay consistent | Automating compliance reports that require exact table grid dimensions | Batch converting marketing dashboards to PDFs without losing border fidelity
// AI Prompts: Write C# code that uses Aspose.Cells to turn an HTML file into a PDF and keep each table border at its original pixel width. | Describe the impact of HtmlSaveOptions.IsBorderCollapsed and ExportSimilarBorderStyle on PDF output of HTML tables in Aspose.Cells. | Propose a method to maintain border widths during HTML‑to‑PDF conversion without writing an intermediate HTML file.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdf
{
    // This example loads an HTML file into an Aspose.Cells Workbook, disables border collapsing (HtmlSaveOptions.IsBorderCollapsed = false) and enables ExportSimilarBorderStyle to retain each border’s pixel thickness. After optionally saving an intermediate HTML file, the workbook is rendered to PDF, producing a document where table borders match the original HTML dimensions.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains the table.
            string htmlPath = "input.html";

            // Load the HTML file into a Workbook.
            // This uses the built‑in constructor that accepts a file name.
            Workbook workbook = new Workbook(htmlPath);

            // Ensure that table borders are not collapsed so that each border is rendered
            // as a separate line with its exact pixel width.
            // The HtmlSaveOptions.IsBorderCollapsed property controls this behavior.
            // We save the workbook back to HTML with this option enabled,
            // then reload it for PDF conversion. This guarantees the border model
            // is preserved before the PDF rendering step.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                IsBorderCollapsed = false,          // keep borders separate
                ExportSimilarBorderStyle = true     // export exact border widths
            };

            // Save the intermediate HTML (optional, can be in memory).
            string intermediateHtml = "intermediate.html";
            workbook.Save(intermediateHtml, htmlOptions);

            // Reload the intermediate HTML to apply the border settings.
            Workbook wbForPdf = new Workbook(intermediateHtml);

            // Convert the workbook to PDF.
            // The Save method with SaveFormat.Pdf uses Aspose.Cells' internal renderer.
            string pdfPath = "output.pdf";
            wbForPdf.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("HTML has been converted to PDF with exact table border widths.");
        }
    }
}
