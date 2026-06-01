using System;
using System.IO;
using Aspose.Cells;

class HtmlToPdfConverter
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Load the HTML file into a Workbook using LoadOptions with LoadFormat.Html
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Configure PDF save options.
            // OnePagePerSheet ensures the whole sheet is rendered on a single PDF page,
            // which helps keep CSS pseudo‑elements (e.g., ::before, ::after) intact.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF file.
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}