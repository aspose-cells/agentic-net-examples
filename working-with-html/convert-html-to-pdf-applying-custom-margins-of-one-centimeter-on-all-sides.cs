// Title: C# – Convert HTML to PDF with 1 cm page margins using Aspose.Cells
// Description: This example shows how to load an HTML file into an Aspose.Cells Workbook, configure the PageSetup object to apply a 1 cm margin on every side, and export the worksheet as a PDF document.
// Keywords: Aspose.Cells | HTML to PDF conversion | C# margin settings | PageSetup centimeters | SaveFormat.Pdf | custom page margins | Aspose.Cells .NET | convert HTML file to PDF | set workbook margins | export worksheet as PDF
// Common Searches: Aspose.Cells set 1 cm margins C# | HTML to PDF conversion with custom margins .NET | C# export HTML as PDF using Aspose.Cells | how to change page margins in Aspose.Cells PDF output | convert multiple HTML files to PDF Aspose.Cells
// Developer Intent: I need to transform an HTML document into a PDF file while enforcing a uniform 1 cm margin on every edge using Aspose.Cells in C#.
// Use Cases: Printing web‑based reports with precise printable area | Generating PDF invoices from HTML templates that require exact margin control | Automating batch conversion of HTML newsletters to PDFs with consistent layout | Creating e‑books from HTML chapters where page margins must meet publishing standards | Preparing legal documents from HTML forms with mandated 1 cm margins
// AI Prompts: Write a C# snippet that reads an HTML file, sets 1 cm top, bottom, left, and right margins via Aspose.Cells PageSetup, and saves it as a PDF. | Explain how to switch the margin unit from centimeters to inches in Aspose.Cells and update the code accordingly. | Provide a C# loop that processes all *.html files in a directory, converts each to PDF, and applies identical 1 cm margins using Aspose.Cells. | Show how to add a header and footer to the PDF while keeping the 1 cm margins. | Demonstrate how to adjust the PDF page size (A4, Letter) together with the custom margins in Aspose.Cells.

using System;
using Aspose.Cells;

namespace HtmlToPdfWithMargins
{
    // This example shows how to load an HTML file into an Aspose.Cells Workbook, configure the PageSetup object to apply a 1 cm margin on every side, and export the worksheet as a PDF document.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlFilePath = "input.html";

            // Path for the resulting PDF file
            string pdfFilePath = "output.pdf";

            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(htmlFilePath);

            // Access the first worksheet (or any specific worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Set custom margins of 1 centimeter on all sides
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.TopMargin = 1.0;      // top margin in centimeters
            pageSetup.BottomMargin = 1.0;   // bottom margin in centimeters
            pageSetup.LeftMargin = 1.0;     // left margin in centimeters
            pageSetup.RightMargin = 1.0;    // right margin in centimeters

            // Save the workbook as a PDF file
            workbook.Save(pdfFilePath, SaveFormat.Pdf);

            Console.WriteLine($"HTML file has been converted to PDF with 1 cm margins: {pdfFilePath}");
        }
    }
}
