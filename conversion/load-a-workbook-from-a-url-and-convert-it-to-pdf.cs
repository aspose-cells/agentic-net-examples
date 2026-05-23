using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells.Utility;

namespace AsposeCellsUrlToPdf
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // URL of the Excel file to be converted
            string excelUrl = "https://example.com/sample.xlsx";

            // Destination PDF file path
            string pdfPath = "output.pdf";

            // Create a temporary file to store the downloaded Excel workbook
            string tempExcelPath = Path.GetTempFileName();

            try
            {
                // Download the Excel file from the URL
                using (HttpClient client = new HttpClient())
                using (HttpResponseMessage response = await client.GetAsync(excelUrl))
                using (Stream contentStream = await response.Content.ReadAsStreamAsync())
                using (FileStream fileStream = new FileStream(tempExcelPath, FileMode.Create, FileAccess.Write))
                {
                    await contentStream.CopyToAsync(fileStream);
                }

                // Convert the downloaded Excel file to PDF using Aspose.Cells ConversionUtility
                ConversionUtility.Convert(tempExcelPath, pdfPath);

                Console.WriteLine($"Conversion completed successfully. PDF saved to: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary Excel file
                if (File.Exists(tempExcelPath))
                {
                    File.Delete(tempExcelPath);
                }
            }
        }
    }
}