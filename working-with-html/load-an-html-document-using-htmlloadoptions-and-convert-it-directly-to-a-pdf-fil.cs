using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFile = "input.html";

        // Desired output PDF file path
        string pdfFile = "output.pdf";

        // Create HtmlLoadOptions to control how the HTML is imported
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.AutoFitColsAndRows = true;   // auto‑fit columns and rows after loading
        loadOptions.SupportDivTag = true;        // enable support for <div> tags

        // Create PDF save options (no specific rule exists, so we instantiate directly)
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // Example: one page per worksheet (optional)
        // saveOptions.OnePagePerSheet = true;

        // Convert the HTML file to PDF using the ConversionUtility method that accepts
        // load and save options. This loads the HTML into a Workbook and saves it as PDF.
        ConversionUtility.Convert(htmlFile, loadOptions, pdfFile, saveOptions);
    }
}