// Title: HTML to PDF conversion using Aspose.Cells for .NET – default settings
// Description: C# example that loads a local HTML file into an Aspose.Cells Workbook with default HtmlLoadOptions and saves it as a PDF using the default SaveFormat.Pdf. A console message confirms the conversion.
// Keywords: Aspose.Cells | C# | .NET | HTML to PDF | HtmlLoadOptions default | SaveFormat.Pdf | server‑side PDF generation | batch HTML conversion | GitHub example | Aspose.Cells HTML PDF
// Common Searches: Aspose.Cells convert HTML to PDF C# | C# load HTML workbook Aspose.Cells | default HtmlLoadOptions Aspose.Cells | save workbook as PDF Aspose | HTML to PDF example Aspose.Cells .NET | how to export HTML as PDF using Aspose.Cells
// Developer Intent: Generate a PDF file from an HTML document with Aspose.Cells using the library’s default configuration.
// Use Cases: Automated nightly conversion of HTML reports to PDF for archival. | Creating PDF invoices from pre‑designed HTML templates in a web service. | Batch processing of marketing assets (HTML) into printable PDF files without custom settings.
// AI Prompts: Provide C# code that converts an HTML file to PDF with Aspose.Cells and adds custom page margins. | Show error‑handling patterns for loading an HTML file and saving it as PDF using Aspose.Cells. | Explain how to convert an HTML string directly to PDF with Aspose.Cells, avoiding intermediate files.

using System;
using Aspose.Cells;               // Aspose.Cells namespace
using Aspose.Cells.Utility;      // For ConversionUtility (not required here but available)

// C# example that loads a local HTML file into an Aspose.Cells Workbook with default HtmlLoadOptions and saves it as a PDF using the default SaveFormat.Pdf. A console message confirms the conversion.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Path for the resulting PDF file
        string pdfFilePath = "output.pdf";

        // Create default HTML load options
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // Load the HTML file into a Workbook using the load options
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Save the workbook as PDF using default save options
        workbook.Save(pdfFilePath, SaveFormat.Pdf);

        Console.WriteLine("HTML file has been successfully converted to PDF.");
    }
}
