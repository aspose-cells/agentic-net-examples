// Title: Convert an HTML file with CSS ::before and ::after pseudo‑elements to PDF using Aspose.Cells in C#
// AI Prompts: Generate C# code that reads an HTML document, loads it into an Aspose.Cells Workbook, and saves it as a PDF while retaining ::before and ::after CSS pseudo‑elements. | Add robust file‑existence validation and exception handling to the HTML‑to‑PDF conversion using Aspose.Cells. | Show how to embed custom fonts in the PDF output when converting HTML that contains CSS pseudo‑elements with Aspose.Cells.
// Common Searches: asp.net convert html with ::before pseudo element to pdf using Aspose.Cells | c# preserve CSS ::after content during html to pdf export with Aspose.Cells | load html and css pseudo‑elements into Aspose.Cells workbook for pdf generation
// Tags: Aspose.Cells HTML to PDF conversion | CSS pseudo‑element support in PDF output | C# Workbook.Load HTML with CSS | Workbook.Save PDF preserving styles | embed custom fonts Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the input HTML file exists, loads it into an Aspose.Cells Workbook (which parses the HTML and associated CSS, including ::before and ::after pseudo‑elements), and saves the workbook as a PDF, with error handling for any conversion issues.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file (must contain CSS with ::before / ::after)
        string htmlFile = "input.html";

        // Desired output PDF file path
        string pdfFile = "output.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlFile))
            {
                Console.WriteLine($"Error: The file '{htmlFile}' was not found.");
                return;
            }

            // Load the HTML document into a Workbook – Aspose.Cells parses HTML and CSS
            var workbook = new Workbook(htmlFile);

            // Save the workbook as PDF
            workbook.Save(pdfFile, SaveFormat.Pdf);

            Console.WriteLine($"Conversion completed. PDF saved to: {pdfFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
