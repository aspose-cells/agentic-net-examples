// Title: Convert HTML to PDF with Aspose.Cells in C# and address ICC color profile embedding limitations
// AI Prompts: Generate C# code that reads an HTML file into an Aspose.Cells Workbook and saves it as a PDF using PdfSaveOptions. | Add validation to confirm the HTML source and ICC profile files exist before conversion, with detailed error handling. | Provide a C# example that uses Aspose.Pdf to embed a custom ICC color profile into a PDF generated from an HTML file.
// Common Searches: how to convert html to pdf with aspose.cells in c# while preserving color accuracy | c# code to embed a custom icc profile into a pdf using Aspose.Pdf after html conversion | why Aspose.Cells PdfSaveOptions cannot embed ICC profiles | checking html and icc files existence before conversion in c#
// Tags: aspose.cells html to pdf c# | pdfsaveoptions icc support limitation | aspose.pdf embed icc profile c# | c# file existence validation for conversion | color accuracy handling in pdf generation

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input HTML and ICC files exist, loads the HTML into an Aspose.Cells Workbook, and saves it as a PDF using PdfSaveOptions. It notes that Aspose.Cells does not currently support embedding ICC profiles directly, suggesting the use of Aspose.Pdf for that purpose, and includes robust exception handling.
class HtmlToPdfWithIcc
{
    static void Main()
    {
        // Paths to the source HTML, destination PDF and the ICC profile file
        string htmlFile = "input.html";
        string pdfFile = "output.pdf";
        string iccFile = "custom.icc";

        try
        {
            // Verify that required input files exist
            if (!File.Exists(htmlFile))
                throw new FileNotFoundException($"HTML source file not found: {htmlFile}");
            if (!File.Exists(iccFile))
                throw new FileNotFoundException($"ICC profile file not found: {iccFile}");

            // Load the HTML document into a Workbook
            Workbook workbook = new Workbook(htmlFile, new HtmlLoadOptions());

            // NOTE: Aspose.Cells PDF save options do not expose a direct ICC profile property.
            // If ICC embedding is required, consider using Aspose.Pdf or a newer version of Aspose.Cells
            // that supports this feature. Here we proceed without embedding the profile.

            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Save the PDF (ICC profile embedding not applied)
            workbook.Save(pdfFile, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
