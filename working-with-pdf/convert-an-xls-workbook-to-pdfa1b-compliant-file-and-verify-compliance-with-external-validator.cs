using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfA1bConversion
{
    // Entry point
    static async Task Main(string[] args)
    {
        // Path to the source Excel file (XLS)
        string excelPath = "input.xls";

        // Path for the resulting PDF/A‑1b file
        string pdfPath = "output_pdfa1b.pdf";

        // Load the workbook from the Excel file (load rule)
        Workbook workbook = new Workbook(excelPath);

        // Create PDF save options and set compliance to PDF/A‑1b (create rule)
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the workbook as PDF/A‑1b (save rule)
        workbook.Save(pdfPath, saveOptions);

        Console.WriteLine($"Workbook successfully converted to PDF/A‑1b: {pdfPath}");

        // Verify the generated PDF with an external validator (free‑form code)
        bool isCompliant = await ValidatePdfA1bAsync(pdfPath);
        Console.WriteLine($"PDF/A‑1b compliance validation result: {isCompliant}");
    }

    // Sends the PDF file to an external validation service and returns the result
    static async Task<bool> ValidatePdfA1bAsync(string pdfFilePath)
    {
        // URL of the external PDF/A‑1b validation service (replace with a real endpoint)
        const string validatorUrl = "https://example.com/api/validate-pdfa1b";

        // Read the PDF file into a byte array
        byte[] pdfBytes = File.ReadAllBytes(pdfFilePath);

        using var httpClient = new HttpClient();

        // Prepare multipart/form-data content with the PDF file
        using var content = new MultipartFormDataContent();
        var fileContent = new ByteArrayContent(pdfBytes);
        fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");
        content.Add(fileContent, "file", Path.GetFileName(pdfFilePath));

        try
        {
            // POST the PDF to the validator
            HttpResponseMessage response = await httpClient.PostAsync(validatorUrl, content);
            response.EnsureSuccessStatusCode();

            // Assume the service returns a JSON with a boolean field "isCompliant"
            string json = await response.Content.ReadAsStringAsync();

            // Simple check: if the response contains "true", consider it compliant
            return json.IndexOf("true", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Validation request failed: {ex.Message}");
            return false;
        }
    }
}