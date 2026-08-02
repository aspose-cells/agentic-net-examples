// Title: Convert HTML Workbook to PDF with Vector‑Rendered Charts using Aspose.Cells for .NET (C#)
// Description: Loads an HTML file into an Aspose.Cells Workbook, saves the workbook as a PDF where charts are rendered as scalable vector graphics, and shows how to export each chart to its own PDF file via Chart.ToPdf.
// Keywords: Aspose.Cells | HTML to PDF conversion | vector chart rendering | Chart.ToPdf | C# | PDFSaveOptions | scalable graphics | workbook to PDF | export chart as PDF | .NET
// Common Searches: Aspose.Cells convert HTML to PDF C# | save workbook as PDF with vector charts | export individual chart to PDF Aspose.Cells | Chart.ToPdf example C# | preserve chart quality when saving PDF
// Developer Intent: Generate PDF output from an HTML workbook while keeping charts as vector graphics.
// Use Cases: Create a single PDF report from an HTML‑based workbook with high‑quality, scalable charts. | Extract each chart from a workbook into separate PDF files for modular reporting. | Automate batch conversion of multiple HTML workbooks to PDFs that retain vector chart rendering.
// AI Prompts: Show how to set custom page size and orientation in PdfSaveOptions while keeping vector chart rendering. | Provide a script to batch‑process a folder of HTML files, converting each to PDF with vector charts using Aspose.Cells. | Explain how to export charts as SVG instead of PDF from the same Workbook instance.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;   // PdfSaveOptions

// Loads an HTML file into an Aspose.Cells Workbook, saves the workbook as a PDF where charts are rendered as scalable vector graphics, and shows how to export each chart to its own PDF file via Chart.ToPdf.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file that contains the worksheet data and charts
        string htmlPath = "input.html";

        // Load the HTML file into a Workbook instance.
        // The Workbook constructor can accept an HTML file and parses it into worksheets.
        Workbook workbook = new Workbook(htmlPath);

        // Configure PDF save options.
        // Vector rendering of charts is the default behavior; no additional settings are required.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the entire workbook as a PDF file.
        workbook.Save("WorkbookOutput.pdf", pdfOptions);

        // Export each chart in the workbook to an individual PDF file.
        // The Chart.ToPdf method renders the chart as vector graphics.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
            {
                Chart chart = sheet.Charts[chartIndex];
                string chartPdfPath = $"Chart_{sheet.Name}_{chartIndex}.pdf";
                chart.ToPdf(chartPdfPath);
            }
        }

        Console.WriteLine("HTML conversion to PDF completed. Vector charts are preserved.");
    }
}
