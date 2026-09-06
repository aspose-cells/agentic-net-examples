// Title: Convert HTML to PDF in C# with Aspose.Cells while preserving line spacing and paragraph indentation
// AI Prompts: Generate C# code that reads an HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, and saves it as a PDF with original whitespace intact. | Demonstrate how to configure HtmlLoadOptions in Aspose.Cells to keep line breaks and indentation when exporting HTML to PDF in a .NET console application. | Add comprehensive error handling to a C# Aspose.Cells HTML‑to‑PDF conversion that checks for missing input files and reports conversion failures.
// Common Searches: aspocells c# retain spacing when converting html to pdf | how to keep paragraph indentation in pdf generated from html using Aspose.Cells | C# console app html to pdf conversion with layout fidelity Aspose.Cells | load html into workbook and export to pdf maintaining original line breaks Aspose.Cells
// Tags: Aspose.Cells HTML import options retain whitespace | HTML to PDF conversion .NET Aspose.Cells | C# load HTML into Workbook for PDF export | keep indentation while exporting HTML to PDF with Aspose.Cells | preserve paragraph formatting in Aspose.Cells PDF output

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// // Reads an HTML file, loads it into an Aspose.Cells Workbook with default HtmlLoadOptions, and saves it as a PDF while preserving the original line breaks and paragraph indentation.
class HtmlToPdfConverter
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the input file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: Input file not found – {htmlPath}");
                return;
            }

            // Read the HTML content preserving original line breaks and indentation
            string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);

            // Load the HTML into an Aspose.Cells Workbook
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                // HtmlLoadOptions can be customized if needed; defaults keep the original layout
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Save the workbook as PDF
                string pdfPath = "output.pdf";
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"HTML has been successfully converted to PDF: {pdfPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
