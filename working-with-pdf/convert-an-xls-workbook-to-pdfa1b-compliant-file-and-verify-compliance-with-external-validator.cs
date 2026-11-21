// Title: Convert an XLS workbook to PDF/A‑1b with Aspose.Cells for .NET and validate the output using an external PDF/A‑1b service
// Common Searches: c# convert xls file to pdf/a-1b using Aspose.Cells | aspnet save workbook as PDF/A-1b with PdfSaveOptions | programmatically validate PDF/A-1b compliance from C# | upload generated PDF to online validator using HttpClient | read compliance flag from JSON response in .NET
// Tags: Aspose.Cells PDF/A-1b save options | C# workbook conversion to PDF/A-1b | HttpClient multipart PDF upload for validation | Parse validator JSON response in .NET | PDF/A-1b compliance verification workflow

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfAExample
{
    // The example loads an XLS workbook with Aspose.Cells, saves it as a PDF/A‑1b file using PdfSaveOptions (Compliance = PdfA1b), then posts the PDF to a configurable external validator via HttpClient and checks the JSON response for a compliance flag to confirm PDF/A‑1b conformance.
    class XlsToPdfAConverter
    {
        // Paths to the source XLS file and the target PDF/A‑1b file
        private readonly string _sourceXlsPath;
        private readonly string _targetPdfPath;
        // URL of an external PDF/A‑1b validator service (expects a file upload and returns JSON)
        private readonly string _validatorUrl;

        public XlsToPdfAConverter(string sourceXlsPath, string targetPdfPath, string validatorUrl)
        {
            _sourceXlsPath = sourceXlsPath;
            _targetPdfPath = targetPdfPath;
            _validatorUrl = validatorUrl;
        }

        public async Task RunAsync()
        {
            try
            {
                // Verify source file exists
                if (!File.Exists(_sourceXlsPath))
                {
                    Console.WriteLine($"Source file not found: {_sourceXlsPath}");
                    return;
                }

                // Load the XLS workbook
                Workbook workbook = new Workbook(_sourceXlsPath);

                // Configure PDF save options for PDF/A‑1b compliance
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1b
                    // EmbedStandardFonts property is not available in current API version; fonts are embedded by default for PDF/A
                };

                // Save the workbook as PDF/A‑1b
                workbook.Save(_targetPdfPath, pdfOptions);
                Console.WriteLine($"PDF/A‑1b file saved to: {_targetPdfPath}");

                // Verify the generated PDF/A‑1b file with the external validator
                bool isCompliant = await VerifyPdfAComplianceAsync(_targetPdfPath);
                Console.WriteLine(isCompliant
                    ? "The PDF/A‑1b file is compliant."
                    : "The PDF/A‑1b file is NOT compliant.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }

        private async Task<bool> VerifyPdfAComplianceAsync(string pdfPath)
        {
            try
            {
                // Ensure the PDF file exists before sending
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine($"PDF file not found for validation: {pdfPath}");
                    return false;
                }

                using (HttpClient client = new HttpClient())
                using (MultipartFormDataContent content = new MultipartFormDataContent())
                using (FileStream fileStream = File.OpenRead(pdfPath))
                {
                    // Add the PDF file to the request
                    var fileContent = new StreamContent(fileStream);
                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/pdf");
                    content.Add(fileContent, "file", Path.GetFileName(pdfPath));

                    // Send POST request to the validator service
                    HttpResponseMessage response = await client.PostAsync(_validatorUrl, content);
                    response.EnsureSuccessStatusCode();

                    // Assume the validator returns JSON: { "compliant": true/false, ... }
                    string json = await response.Content.ReadAsStringAsync();

                    // Simple check for the "compliant" flag (real implementation would parse JSON properly)
                    return json.IndexOf("\"compliant\":true", StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Validation error: {ex.Message}");
                return false;
            }
        }

        // Example usage
        static async Task Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: XlsToPdfAConverter <source.xls> <target.pdf> <validatorUrl>");
                return;
            }

            var converter = new XlsToPdfAConverter(args[0], args[1], args[2]);
            await converter.RunAsync();
        }
    }
}
