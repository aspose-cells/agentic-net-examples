// Title: C# – Convert an XLS workbook to PDF/A‑1b with Aspose.Cells and validate the result via a REST validator
// AI Prompts: Write C# code that loads an .xls file using Aspose.Cells, configures PdfSaveOptions for PDF/A‑1b compliance, and saves the workbook as a PDF/A‑1b document. | Create a C# method that sends a generated PDF/A‑1b file to a remote validation service with HttpClient multipart/form-data and parses the JSON response to determine compliance. | Show error‑handling logic for missing source files and non‑successful responses from the PDF/A validation endpoint.
// Common Searches: aspnet convert excel xls to pdf/a-1b using aspose.cells | c# upload pdf/a-1b file to validation api and read json result | how to set PdfSaveOptions.Compliance to PdfA1b in Aspose.Cells | verify pdf/a-1b compliance after converting excel workbook | sample code for posting a file to external validator with HttpClient in .NET
// Tags: Aspose.Cells PDF/A‑1b conversion C# | PdfSaveOptions compliance PDF/A‑1b | HttpClient multipart file upload .NET | PDF/A validation API integration | Excel to PDF/A conversion example

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an .xls workbook with Aspose.Cells, saves it as a PDF/A‑1b file by setting PdfSaveOptions.Compliance, then uploads the PDF to an external PDF/A‑1b validator using HttpClient multipart/form-data and interprets the JSON response to confirm compliance.
class PdfAConversion
{
    // Path to the source XLS workbook
    private const string SourceXlsPath = @"C:\Input\workbook.xls";

    // Path where the PDF/A‑1b file will be saved
    private const string OutputPdfPath = @"C:\Output\workbook_pdfa1b.pdf";

    // URL of an external PDF/A‑1b validator service (example placeholder)
    private const string ValidatorUrl = "https://api.pdfavalidator.example.com/validate";

    static async Task Main()
    {
        try
        {
            // Verify source workbook exists
            if (!File.Exists(SourceXlsPath))
            {
                Console.WriteLine($"Source workbook not found: {SourceXlsPath}");
                return;
            }

            // Load the XLS workbook
            Workbook workbook = new Workbook(SourceXlsPath);

            // Configure PDF save options for PDF/A‑1b compliance
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                // Set the compliance level to PDF/A‑1b
                Compliance = PdfCompliance.PdfA1b
                // Note: EmbedStandardFonts is not required / not available in this version
            };

            // Save the workbook as a PDF/A‑1b file
            workbook.Save(OutputPdfPath, saveOptions);
            Console.WriteLine($"Workbook converted and saved to PDF/A‑1b at: {OutputPdfPath}");

            // Verify the generated PDF/A‑1b file using an external validator
            bool isCompliant = await VerifyPdfAComplianceAsync(OutputPdfPath);
            Console.WriteLine(isCompliant
                ? "The PDF/A‑1b file is compliant."
                : "The PDF/A‑1b file is NOT compliant.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="pdfPath">Full path to the PDF/A‑1b file.</param>
    /// <returns>True if the file is compliant; otherwise false.</returns>
    private static async Task<bool> VerifyPdfAComplianceAsync(string pdfPath)
    {
        if (!File.Exists(pdfPath))
        {
            Console.WriteLine("PDF file not found for validation.");
            return false;
        }

        using (HttpClient client = new HttpClient())
        using (MultipartFormDataContent content = new MultipartFormDataContent())
        using (FileStream fileStream = File.OpenRead(pdfPath))
        {
            // Add the PDF file to the multipart content
            StreamContent fileContent = new StreamContent(fileStream);
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/pdf");
            content.Add(fileContent, "file", Path.GetFileName(pdfPath));

            try
            {
                // POST the file to the validator endpoint
                HttpResponseMessage response = await client.PostAsync(ValidatorUrl, content);
                response.EnsureSuccessStatusCode();

                // Assume the validator returns JSON with a boolean field "isCompliant"
                string jsonResponse = await response.Content.ReadAsStringAsync();

                // Simple parsing (replace with proper JSON deserialization as needed)
                return jsonResponse.Contains("\"isCompliant\":true");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Validation request failed: {ex.Message}");
                return false;
            }
        }
    }
}
