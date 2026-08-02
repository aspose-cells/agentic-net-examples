// Title: Convert HTML to PDF (A4 Landscape) with Aspose.Cells for .NET
// Description: Load an HTML file into an Aspose.Cells Workbook, set the worksheet's page setup to A4 size and landscape orientation, and export the result as a PDF using C#.
// Keywords: Aspose.Cells HTML to PDF | C# convert HTML to PDF | A4 landscape PDF Aspose.Cells | set page orientation Aspose.Cells | Workbook SaveFormat.Pdf | Aspose.Cells page setup
// Common Searches: Aspose.Cells convert HTML file to PDF A4 landscape C# | how to set paper size and orientation before saving PDF in Aspose.Cells | C# export HTML workbook to PDF with landscape layout | Aspose.Cells page setup A4 landscape example
// Developer Intent: Generate a PDF from an HTML document and configure the output to A4 landscape using Aspose.Cells for .NET.
// Use Cases: Transform web‑based reports saved as HTML into printable A4‑landscape PDFs. | Batch‑process HTML invoices, applying A4 landscape settings before PDF export. | Create brochure‑style PDFs from HTML templates with a consistent landscape layout.
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook, sets A4 paper size and landscape orientation, and saves it as a PDF. | Explain how to adjust margins, scaling, and header/footer settings when converting HTML to A4 landscape PDF with Aspose.Cells. | Show a loop that processes multiple HTML files, applies A4 landscape page setup to each worksheet, and exports each to a separate PDF.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    // Load an HTML file into an Aspose.Cells Workbook, set the worksheet's page setup to A4 size and landscape orientation, and export the result as a PDF using C#.
    class Program
    {
        static void Main()
        {
            // Load the HTML file into a workbook.
            // The Workbook constructor can accept an HTML file path.
            Workbook workbook = new Workbook("input.html");

            // Access the first worksheet (or iterate through all worksheets if needed).
            Worksheet sheet = workbook.Worksheets[0];

            // Set the paper size to A4.
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Set the page orientation to Landscape.
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Save the workbook as a PDF file.
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
