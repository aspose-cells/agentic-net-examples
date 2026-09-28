// Title: Convert an Excel workbook downloaded from a URL to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that retrieves an .xlsx file from a remote URL with HttpClient, loads it into an Aspose.Cells Workbook via a MemoryStream, and exports it as a PDF. | Add error‑handling that falls back to a local .xlsx file when the HTTP download fails, then performs the same PDF conversion. | Modify the program to accept the source URL and destination PDF path as command‑line arguments while keeping the download‑fallback logic.
// Common Searches: aspnet convert excel from http url to pdf using aspose.cells | c# download xlsx to memory stream and save as pdf with aspose | fallback to local excel file if remote download fails aspose.cells example | load workbook from stream and export to pdf in .net core | aspose.cells remote file conversion to pdf command line
// Tags: remote Excel to PDF conversion Aspose.Cells | stream-based workbook loading C# | fallback mechanism for file download Aspose.Cells | save workbook as PDF Aspose.Cells | download Excel via HttpClient MemoryStream

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// Downloads an Excel file from a specified URL (with a local fallback), loads it into an Aspose.Cells Workbook, and saves the workbook as a PDF document.
public class UrlToPdfConverter
{
    // Entry point
    public static async Task Main()
    {
        // URL of the Excel file to be converted
        string excelUrl = "https://example.com/sample.xlsx";

        // Local fallback path for the Excel file (if download fails)
        string localExcelPath = "sample.xlsx";

        // Local path for the resulting PDF
        string pdfPath = "output.pdf";

        try
        {
            // Attempt to download the Excel file into a memory stream
            using (MemoryStream excelStream = await DownloadFileAsync(excelUrl))
            {
                Workbook workbook;

                if (excelStream != null && excelStream.Length > 0)
                {
                    // Load the workbook from the downloaded stream
                    workbook = new Workbook(excelStream);
                }
                else if (File.Exists(localExcelPath))
                {
                    // Fallback: load workbook from a local file
                    workbook = new Workbook(localExcelPath);
                }
                else
                {
                    throw new FileNotFoundException("Neither the remote Excel file could be downloaded nor the local fallback file was found.");
                }

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Conversion completed: {pdfPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }

    // Helper method to download a file into a MemoryStream
    private static async Task<MemoryStream> DownloadFileAsync(string url)
    {
        try
        {
            using (HttpClient client = new HttpClient())
            {
                byte[] data = await client.GetByteArrayAsync(url);
                var stream = new MemoryStream(data);
                // Ensure the stream position is at the beginning
                stream.Position = 0;
                return stream;
            }
        }
        catch (HttpRequestException httpEx)
        {
            // Log the HTTP error and return null to trigger fallback logic
            Console.WriteLine($"Failed to download file from URL: {httpEx.Message}");
            return null;
        }
        catch (Exception ex)
        {
            // Log any other errors and return null
            Console.WriteLine($"Unexpected error during download: {ex.Message}");
            return null;
        }
    }
}
