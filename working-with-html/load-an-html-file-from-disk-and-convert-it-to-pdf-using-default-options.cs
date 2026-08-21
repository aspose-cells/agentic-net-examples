// Title: Convert a local HTML file to PDF with Aspose.Cells for .NET (default settings)
// Description: A concise C# example that reads an HTML file from disk and uses Aspose.Cells.Utility.ConversionUtility.Convert to generate a PDF with the library's built‑in load and save options, then confirms success on the console.
// Keywords: Aspose.Cells HTML to PDF | C# ConversionUtility | load HTML file .NET | export PDF default options | Aspose.Cells Utility Convert example | HTML to PDF conversion .NET | Aspose.Cells PDF export
// Common Searches: Aspose.Cells convert HTML file to PDF C# | ConversionUtility default HTML to PDF example | How to export HTML as PDF using Aspose.Cells | C# code for HTML to PDF with Aspose.Cells
// Developer Intent: Generate a PDF from a local HTML document using Aspose.Cells with no custom configuration.
// Use Cases: Archive a batch of HTML reports as PDFs for compliance. | Create printable invoices from HTML templates in a .NET service. | Automate nightly conversion of web‑based documentation to PDF for distribution.
// AI Prompts: Show C# code that converts an HTML string to PDF with custom margins using Aspose.Cells. | Add robust error handling around ConversionUtility.Convert for missing or malformed HTML files. | Write a script that scans a folder for *.html files and converts each to a matching PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A concise C# example that reads an HTML file from disk and uses Aspose.Cells.Utility.ConversionUtility.Convert to generate a PDF with the library's built‑in load and save options, then confirms success on the console.
class Program
{
    static void Main()
    {
        // Path to the HTML file that will be converted
        string htmlPath = "input.html";

        // Desired output PDF file path
        string pdfPath = "output.pdf";

        // Convert the HTML file to PDF using default load and save options
        // This utilizes the provided ConversionUtility.Convert method
        ConversionUtility.Convert(htmlPath, pdfPath);

        Console.WriteLine("HTML file has been successfully converted to PDF.");
    }
}
