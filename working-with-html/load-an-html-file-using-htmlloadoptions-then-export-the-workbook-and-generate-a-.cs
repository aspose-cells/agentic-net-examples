using System;
using Aspose.Cells;

namespace HtmlToPdfConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Configure HTML load options to preserve layout
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                // Enable support for <div> tags which affect layout
                SupportDivTag = true,
                // Auto‑fit rows and columns based on the HTML content
                AutoFitColsAndRows = true,
                // Load formulas if the HTML contains them
                LoadFormulas = true
            };

            // Load the HTML file into a workbook using the configured options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            // Save the workbook as PDF, preserving the layout from the HTML
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF '{pdfPath}'.");
        }
    }
}