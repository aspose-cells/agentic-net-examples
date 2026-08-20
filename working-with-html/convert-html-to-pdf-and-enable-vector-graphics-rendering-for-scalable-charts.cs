// Title: C# – Convert HTML to PDF with Aspose.Cells while keeping charts as vector graphics (PDF/A‑1b)
// Description: This example shows how to load an HTML file into an Aspose.Cells Workbook, optionally add a chart, and save the workbook as a PDF. Charts and shapes are rendered automatically as scalable vector graphics, and PDF/A‑1b compliance can be enabled via PdfSaveOptions.
// Keywords: Aspose.Cells HTML to PDF C# | vector chart PDF conversion | PdfSaveOptions PDF/A‑1b | scalable graphics Aspose.Cells | HTML workbook to PDF example
// Common Searches: Aspose.Cells convert HTML to PDF C# | how to keep charts vector when saving PDF with Aspose.Cells | PDF/A‑1b output from HTML workbook Aspose | C# code for HTML to PDF with vector graphics Aspose.Cells
// Developer Intent: Create a PDF from an HTML workbook and ensure all charts are rendered as scalable vector graphics, optionally complying with PDF/A‑1b.
// Use Cases: Generate print‑ready PDFs from marketing HTML reports that contain charts. | Produce archival PDF/A‑1b documents from web‑based dashboards while preserving chart clarity. | Automate batch conversion of HTML files to PDFs with vector‑based charts for scalable distribution.
// AI Prompts: Provide C# code that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF with vector‑rendered charts. | Show how to configure PdfSaveOptions for PDF/A‑1b compliance when converting HTML to PDF using Aspose.Cells. | Explain why Aspose.Cells renders charts as vector graphics in the resulting PDF and how this benefits scalability.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// This example shows how to load an HTML file into an Aspose.Cells Workbook, optionally add a chart, and save the workbook as a PDF. Charts and shapes are rendered automatically as scalable vector graphics, and PDF/A‑1b compliance can be enabled via PdfSaveOptions.
class HtmlToPdfVectorDemo
{
    static void Main()
    {
        // Load the HTML file into a workbook.
        // The Workbook constructor can accept a file path and will parse the HTML content.
        Workbook workbook = new Workbook("input.html");

        // ------------------------------------------------------------
        // OPTIONAL: Add a sample chart to demonstrate that charts are
        // rendered as vector graphics when the workbook is saved to PDF.
        // ------------------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet contains data for the chart.
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);

        // Add a column chart.
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);               // Values
        chart.NSeries.CategoryData = "A2:A3";           // Categories
        chart.Title.Text = "Sample Chart";

        // ------------------------------------------------------------
        // Configure PDF save options.
        // Charts and shapes are always rendered as vector elements,
        // so no additional settings are required for scalability.
        // ------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // (Optional) Set PDF/A compliance if needed.
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the workbook as a PDF file.
        workbook.Save("output.pdf", pdfOptions);
    }
}
