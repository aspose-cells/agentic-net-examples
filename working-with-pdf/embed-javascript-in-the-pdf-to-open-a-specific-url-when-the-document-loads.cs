using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add a hyperlink to a cell – this demonstrates hyperlink usage
        worksheet.Hyperlinks.Add("A1", 1, 1, "http://example.com");
        worksheet.Hyperlinks[0].TextToDisplay = "Open Example";

        // ------------------------------------------------------------
        // NOTE: Aspose.Cells does not expose a documented API for embedding
        // JavaScript into a PDF (e.g., app.launchURL) in the current version.
        // The following placeholder shows where such code would be placed
        // once the appropriate property/method is available.
        // ------------------------------------------------------------
        // PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // pdfOptions.JavaScript = "app.launchURL('http://example.com', true);";
        // workbook.Save("Output.pdf", pdfOptions);

        // Fallback: save PDF without JavaScript
        workbook.Save("Output.pdf");
    }
}