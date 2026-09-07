// Title: Convert HTML with linked external CSS to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an HTML file with external CSS files into an Aspose.Cells Workbook using HtmlLoadOptions and saves it as a PDF. | Demonstrate how to configure Aspose.Cells to automatically include linked stylesheet rules when converting HTML to PDF in a .NET application. | Provide error‑handling logic for missing HTML input files and exceptions during HTML‑to‑PDF conversion with Aspose.Cells.
// Common Searches: asp.net core convert html page that references external css to pdf with aspose.cells | c# sample that loads html and linked stylesheet then exports to pdf using aspose.cells | how to ensure linked css styles are applied when saving html as pdf in aspose.cells | example of checking html file existence before workbook load in aspose.cells conversion
// Tags: CSS file inclusion in Aspose.Cells HTML load | preserve stylesheet formatting during PDF generation | console app for styled HTML conversion | pre‑load file existence verification for workbook creation | SaveFormat.Pdf usage with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the input HTML file exists, loads it into an Aspose.Cells Workbook with default HtmlLoadOptions (which automatically processes linked CSS), and saves the result as a PDF using SaveFormat.Pdf, while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Path for the generated PDF file
            string pdfPath = "output.pdf";

            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            // Configure loading options (default loads external CSS and images)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a new workbook
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as PDF, preserving the original stylesheet formatting
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
