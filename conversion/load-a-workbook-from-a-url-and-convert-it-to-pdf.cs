// Title: Download an Excel file from a URL and convert it to PDF with Aspose.Cells for .NET (C#)
// Description: This example shows how to use HttpClient to retrieve an .xlsx workbook from a remote URL, load it into an Aspose.Cells Workbook via a MemoryStream, and save the workbook directly as a PDF file. It includes basic error handling for network and file issues.
// Keywords: Aspose.Cells | C# | download Excel from URL | convert Excel to PDF | MemoryStream | HttpClient | PDF conversion .NET | remote workbook loading | Aspose.Cells SaveFormat.Pdf
// Common Searches: How to load an Excel file from a web address using Aspose.Cells | Convert remote .xlsx to PDF in C# | Aspose.Cells example for downloading and converting Excel | Save Aspose.Cells workbook as PDF without saving the source file
// Developer Intent: Retrieve an Excel workbook from a web URL and generate a PDF version using Aspose.Cells in C#.
// Use Cases: Create PDF reports from Excel files hosted on public APIs. | Process incoming Excel data in memory and produce PDF invoices without persisting the original file. | Build a web service that returns a PDF rendering of Excel documents stored in cloud storage.
// AI Prompts: Write C# code that downloads an .xlsx file from a given URL, loads it into an Aspose.Cells Workbook via MemoryStream, and saves it as a PDF. | Explain how to handle HTTP errors and reset the MemoryStream before opening the workbook with Aspose.Cells. | Show how to extend the sample to export the workbook to other formats such as PNG, DOCX, or HTML using Aspose.Cells.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ConvertWorkbook
{
    // This example shows how to use HttpClient to retrieve an .xlsx workbook from a remote URL, load it into an Aspose.Cells Workbook via a MemoryStream, and save the workbook directly as a PDF file. It includes basic error handling for network and file issues.
    class ConvertWorkbookFromUrlToPdf
    {
        static async Task Main(string[] args)
        {
            // URL of the Excel file to be downloaded
            string fileUrl = "https://example.com/sample.xlsx";

            // Local path for the resulting PDF file
            string pdfPath = "output.pdf";

            try
            {
                using (HttpClient httpClient = new HttpClient())
                {
                    // Get the response and ensure it succeeded
                    HttpResponseMessage response = await httpClient.GetAsync(fileUrl);
                    response.EnsureSuccessStatusCode();

                    // Read the Excel content into a seekable memory stream
                    using (Stream excelStream = await response.Content.ReadAsStreamAsync())
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        await excelStream.CopyToAsync(memoryStream);
                        memoryStream.Position = 0; // Reset for reading

                        // Load the workbook from the memory stream
                        Workbook workbook = new Workbook(memoryStream);

                        // Save the workbook as PDF
                        workbook.Save(pdfPath, SaveFormat.Pdf);
                    }
                }

                Console.WriteLine($"Conversion completed. PDF saved to: {pdfPath}");
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Error downloading the Excel file: {ex.Message}");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File not found: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
