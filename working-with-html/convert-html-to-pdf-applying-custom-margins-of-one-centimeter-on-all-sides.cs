// Title: C# – Convert HTML to PDF with 1 cm margins using Aspose.Cells
// Description: This C# example loads an HTML file into an Aspose.Cells Workbook, sets all page margins to 1 centimeter via the PageSetup object, and saves the workbook as a PDF. It demonstrates precise margin control for printable output.
// Keywords: Aspose.Cells | HTML to PDF | C# margin settings | page setup centimeters | custom PDF margins | Aspose.Cells SaveFormat.Pdf | convert HTML file to PDF | set page margins programmatically
// Common Searches: Aspose.Cells set 1 cm margins | C# convert HTML to PDF with custom margins | How to change page margins in Aspose.Cells PDF export | HTML to PDF conversion with centimeter margins .NET | Aspose.Cells page setup example
// Developer Intent: Create a PDF from an HTML source and enforce a uniform 1 cm margin on every side using Aspose.Cells in C#.
// Use Cases: Generating printable invoices from web‑based HTML templates with consistent margin requirements. | Batch‑processing marketing brochures where a 1 cm whitespace border meets publishing standards. | Automating report generation from dashboards, ensuring PDFs have exact margin dimensions for binding.
// AI Prompts: Provide C# code that converts an HTML file to PDF with 2 cm margins using Aspose.Cells. | Explain how to switch between portrait and landscape orientation while keeping centimeter‑based margins in Aspose.Cells PDF export. | Show how to add error handling for missing HTML files and invalid margin values during conversion with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    // This C# example loads an HTML file into an Aspose.Cells Workbook, sets all page margins to 1 centimeter via the PageSetup object, and saves the workbook as a PDF. It demonstrates precise margin control for printable output.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(htmlPath);

            // Access the first worksheet (the HTML content is loaded into the first sheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Set custom margins of 1 centimeter on all sides
            worksheet.PageSetup.TopMargin = 1.0;      // top margin (cm)
            worksheet.PageSetup.BottomMargin = 1.0;   // bottom margin (cm)
            worksheet.PageSetup.LeftMargin = 1.0;     // left margin (cm)
            worksheet.PageSetup.RightMargin = 1.0;    // right margin (cm)

            // Save the workbook as a PDF file
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF with 1 cm margins and saved as '{pdfPath}'.");
        }
    }
}
