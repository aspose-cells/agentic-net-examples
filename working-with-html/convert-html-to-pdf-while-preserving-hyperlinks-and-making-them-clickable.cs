// Title: Convert an HTML file to a PDF with active hyperlinks using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an HTML document into an Aspose.Cells Workbook with LoadOptions set to LoadFormat.Html and saves it as a PDF preserving clickable hyperlinks. | Show how to configure Aspose.Cells to retain anchor‑tag links when exporting an HTML workbook to PDF in a .NET application. | Provide a minimal example that converts input.html to output.pdf using Aspose.Cells, ensuring all hyperlinks remain functional in the generated PDF.
// Common Searches: Aspose.Cells C# convert HTML to PDF while keeping hyperlinks clickable | how to preserve anchor tags when exporting HTML to PDF with Aspose.Cells .NET | sample code for loading HTML into Workbook and saving as PDF with active links in C#
// Tags: load HTML into Aspose.Cells Workbook | save workbook as PDF with hyperlinks | LoadOptions LoadFormat.Html usage | HTML to PDF conversion Aspose.Cells .NET | clickable hyperlink retention in PDF export

using System;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells Workbook using LoadOptions (LoadFormat.Html) and saves it as a PDF, preserving all hyperlinks as clickable links in the output document.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path for the resulting PDF file
        string pdfPath = "output.pdf";

        // Load the HTML file into a Workbook.
        // LoadOptions with LoadFormat.Html tells Aspose.Cells to interpret the file as HTML.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Save the workbook as PDF.
        // Hyperlinks present in the HTML are converted to cell hyperlinks
        // and are retained as clickable links in the generated PDF.
        workbook.Save(pdfPath, SaveFormat.Pdf);
    }
}
