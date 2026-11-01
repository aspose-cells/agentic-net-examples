// Title: Convert an XLS workbook to PDF/A‑1b with Aspose.Cells for .NET and verify compliance using a REST validator
// AI Prompts: Write C# code that loads an .xls file with Aspose.Cells, sets PdfSaveOptions.Compliance to PdfA1b, and saves the workbook as a PDF/A‑1b document. | Create a method that posts a PDF file to a validator endpoint using HttpClient multipart/form-data and returns true when the JSON response contains "compliant":true. | Add robust error handling to detect missing source or output files, catch conversion and validation exceptions, and log clear status messages.
// Common Searches: how to save Excel workbook as PDF/A-1b using Aspose.Cells in C# | c# program to validate PDF/A compliance with external web service | Aspose.Cells convert xls to PDF/A-1b and check compliance | post PDF file to REST validator endpoint using HttpClient in .NET | verify PDF/A-1b compliance after converting from Excel
// Tags: Aspose.Cells PDF/A-1b export C# | PdfSaveOptions compliance configuration | C# multipart/form-data PDF upload | REST API PDF/A validation service | XLS to PDF/A-1b conversion workflow

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an XLS workbook, saves it as a PDF/A‑1b file using Aspose.Cells PdfSaveOptions, uploads the PDF to a REST validator via HttpClient multipart request, and reports whether the file meets PDF/A‑1b compliance.
class Program
{
    // Entry point
    static async Task Main(string[] args)
    {
        try
        {
            // Path to the source XLS workbook
            string sourcePath = "input.xls";

            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Path for the generated PDF/A‑1b file
            string pdfPath = "output.pdf";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options for PDF/A‑1b compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set PDF/A‑1b compliance
                Compliance = PdfCompliance.PdfA1b
                // Additional options such as embedding fonts or setting document title
                // are not available directly via PdfSaveOptions in Aspose.Cells.
            };

            // Save the workbook as PDF/A‑1b
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF/A‑1b file saved to: {pdfPath}");

            // Verify compliance using an external validator service
            // (Replace the URL with the actual validator endpoint)
            string validatorUrl = "https://example-validator.com/api/validatePdfA";

            // Ensure the generated PDF exists before validation
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"Generated PDF not found: {pdfPath}");
                return;
            }

            // Call the validator and output the result
            bool isCompliant = await ValidatePdfAAsync(pdfPath, validatorUrl);
            Console.WriteLine(isCompliant
                ? "The PDF/A‑1b file is compliant."
                : "The PDF/A‑1b file is NOT compliant.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Sends the PDF file to an external validator and returns true if compliant
    private static async Task<bool> ValidatePdfAAsync(string pdfFilePath, string validatorEndpoint)
    {
        try
        {
            using (HttpClient client = new HttpClient())
            using (MultipartFormDataContent content = new MultipartFormDataContent())
            using (FileStream fileStream = File.OpenRead(pdfFilePath))
            {
                // Add the PDF file to the request
                var fileContent = new StreamContent(fileStream);
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/pdf");
                content.Add(fileContent, "file", Path.GetFileName(pdfFilePath));

                // POST the file to the validator service
                HttpResponseMessage response = await client.PostAsync(validatorEndpoint, content);
                response.EnsureSuccessStatusCode();

                // Assume the validator returns JSON like { "compliant": true }
                string json = await response.Content.ReadAsStringAsync();

                // Simple check for "true" in the response (adjust parsing as needed)
                return json.Contains("\"compliant\":true");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Validation request failed: {ex.Message}");
            return false;
        }
    }
}
