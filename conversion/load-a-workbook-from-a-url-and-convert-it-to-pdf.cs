// Title: Download an Excel workbook from a URL and convert it to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# download .xlsx from a web address and save as PDF with Aspose.Cells | Aspose.Cells load workbook from HttpClient stream and export to PDF | How to handle download failure and use a local Excel file as fallback in Aspose.Cells conversion | Convert remote Excel file to PDF in .NET without saving the file locally
// Tags: download Excel with HttpClient Aspose.Cells | load workbook from stream Aspose.Cells C# | export workbook to PDF SaveFormat.Pdf | fallback local Excel loading Aspose.Cells | exception handling for Excel download Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example demonstrates how to fetch an Excel file from a remote URL using HttpClient, load it into an Aspose.Cells Workbook via a stream, and save the workbook as a PDF. It also includes a fallback mechanism that loads a local .xlsx file if the download fails, with comprehensive exception handling for network, loading, and saving errors.
class Program
{
    static async Task Main(string[] args)
    {
        // URL of the Excel file to be converted
        string excelUrl = "https://example.com/sample.xlsx";

        // Destination PDF file path
        string pdfPath = "output.pdf";

        Workbook workbook = null;

        try
        {
            // Attempt to download the Excel file
            using (HttpClient httpClient = new HttpClient())
            using (Stream excelStream = await httpClient.GetStreamAsync(excelUrl))
            {
                // Load the workbook from the downloaded stream
                workbook = new Workbook(excelStream);
            }
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"Failed to download Excel file: {ex.Message}");
            // Fallback to a local template file if it exists
            string localPath = "sample.xlsx";
            if (File.Exists(localPath))
            {
                try
                {
                    workbook = new Workbook(localPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load local workbook: {loadEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Local fallback file not found: {localPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }

        if (workbook != null)
        {
            try
            {
                // Convert and save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving PDF: {saveEx.Message}");
            }
        }
        else
        {
            Console.WriteLine("No workbook available for conversion.");
        }
    }
}
