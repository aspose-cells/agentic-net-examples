// Title: C# – Convert XLS to PDF/A‑1b with Aspose.Cells and Validate via Web Service
// Description: Sample C# console app that loads or creates an XLS workbook, uses Aspose.Cells PdfSaveOptions (Compliance = PdfA1b) to generate a PDF/A‑1b file, then posts the PDF to a remote validation API and parses the JSON result to confirm compliance.
// Keywords: Aspose.Cells | C# PDF/A-1b conversion | XLS to PDF/A-1b | PdfSaveOptions PdfA1b | .NET PDF/A validation | external PDF/A validator | archival PDF compliance | Aspose.Cells PDF/A example
// Common Searches: how to save Excel as PDF/A-1b using Aspose.Cells C# | C# code to check PDF/A-1b compliance with a web service | Aspose.Cells PdfSaveOptions PDF/A-1b sample | validate generated PDF/A-1b file programmatically | convert legacy XLS reports to PDF/A for long‑term storage
// Developer Intent: Create a PDF/A‑1b compliant document from an XLS workbook and programmatically confirm its compliance through an external validation service.
// Use Cases: Automated archival of legacy Excel reports by converting them to PDF/A‑1b for regulatory long‑term storage. | CI/CD pipeline step that generates PDFs from spreadsheets and fails the build if PDF/A‑1b compliance is not met. | Web API that accepts XLS uploads, returns a PDF/A‑1b version and a compliance flag from a third‑party validator. | Batch processing of financial statements to ensure PDF/A‑1b compliance before distribution to auditors.
// AI Prompts: Generate C# code using Aspose.Cells to convert an Excel file to PDF/A‑1b and log any conversion warnings. | Write a method that uploads a PDF/A‑1b file to a RESTful validator, handles network errors, and returns a boolean compliance result without external JSON libraries. | Suggest performance optimizations for converting large workbooks to PDF/A‑1b and performing asynchronous validation calls.

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfA1bConversion
{
    // Sample C# console app that loads or creates an XLS workbook, uses Aspose.Cells PdfSaveOptions (Compliance = PdfA1b) to generate a PDF/A‑1b file, then posts the PDF to a remote validation API and parses the JSON result to confirm compliance.
    class Program
    {
        // Entry point
        static async Task Main(string[] args)
        {
            try
            {
                // Paths for source Excel file and destination PDF/A‑1b file
                string sourceExcelPath = "input.xls";
                string destinationPdfPath = "output_pdfa1b.pdf";

                // Ensure the source Excel file exists; create a simple one if missing
                if (!File.Exists(sourceExcelPath))
                {
                    var tempWb = new Workbook();
                    tempWb.Worksheets[0].Cells["A1"].PutValue("Sample data");
                    tempWb.Save(sourceExcelPath);
                }

                // Load the workbook from the Excel file
                Workbook workbook = new Workbook(sourceExcelPath);

                // Configure PDF save options for PDF/A‑1b compliance
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1b
                };

                // Save the workbook as a PDF/A‑1b file
                workbook.Save(destinationPdfPath, pdfOptions);
                Console.WriteLine($"Workbook saved as PDF/A‑1b to '{destinationPdfPath}'.");

                // Verify the generated PDF with an external validator
                // (Replace the URL with the actual validator endpoint)
                string validatorUrl = "https://example.com/api/validatePdfA1b";

                bool isCompliant = await VerifyPdfComplianceAsync(destinationPdfPath, validatorUrl);
                Console.WriteLine(isCompliant
                    ? "The PDF is compliant with PDF/A‑1b."
                    : "The PDF is NOT compliant with PDF/A‑1b.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Sends the PDF file to an external validation service and returns the result
        private static async Task<bool> VerifyPdfComplianceAsync(string pdfFilePath, string validatorEndpoint)
        {
            if (!File.Exists(pdfFilePath))
                throw new FileNotFoundException("PDF file not found.", pdfFilePath);

            try
            {
                using (HttpClient client = new HttpClient())
                using (MultipartFormDataContent content = new MultipartFormDataContent())
                using (FileStream pdfStream = File.OpenRead(pdfFilePath))
                {
                    // Add the PDF file to the multipart content
                    StreamContent fileContent = new StreamContent(pdfStream);
                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/pdf");
                    content.Add(fileContent, "file", Path.GetFileName(pdfFilePath));

                    // Post the request to the validator
                    HttpResponseMessage response = await client.PostAsync(validatorEndpoint, content);

                    // If the service does not support POST, treat as non‑compliant
                    if (!response.IsSuccessStatusCode)
                        return false;

                    // Assume the validator returns JSON: { "compliant": true/false }
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    // Simple parsing without external libraries
                    bool compliant = jsonResponse.IndexOf("\"compliant\":true", StringComparison.OrdinalIgnoreCase) >= 0;
                    return compliant;
                }
            }
            catch (HttpRequestException)
            {
                // Network or protocol errors – treat as non‑compliant
                return false;
            }
        }
    }
}
