// Title: Download an Excel file from a URL and convert it to PDF with Aspose.Cells in C#
// Common Searches: C# code to fetch an .xlsx from a web address and save it as PDF using Aspose.Cells | How to load an Excel workbook from a byte array and export to PDF in .NET | Aspose.Cells example for converting a remote Excel file to PDF without writing it to disk | Using WebClient to download an Excel file and convert to PDF with Aspose.Cells | MemoryStream workbook loading and PDF saving with Aspose.Cells in C#
// Tags: download excel file via WebClient Aspose.Cells | load workbook from MemoryStream Aspose.Cells | convert workbook to PDF Aspose.Cells | remote xlsx to pdf conversion .NET | fallback to local file Aspose.Cells conversion

using System;
using System.IO;
using System.Net;
using Aspose.Cells;

// The sample downloads an Excel workbook from a specified URL using WebClient, falls back to a local file if the download fails, loads the file bytes into a MemoryStream, creates an Aspose.Cells Workbook from the stream, and saves the workbook as a PDF document.
class Program
{
    static void Main()
    {
        // URL of the Excel file to be converted
        string excelUrl = "https://example.com/sample.xlsx";

        // Destination PDF file
        string pdfPath = "output.pdf";

        // Try to download the Excel file
        byte[] fileBytes = null;
        try
        {
            using (WebClient client = new WebClient())
            {
                fileBytes = client.DownloadData(excelUrl);
            }
        }
        catch (WebException ex)
        {
            Console.WriteLine($"Failed to download file from URL: {ex.Message}");
        }

        // If download failed, try to load a local template file
        if (fileBytes == null)
        {
            string localPath = "sample.xlsx";
            if (File.Exists(localPath))
            {
                try
                {
                    fileBytes = File.ReadAllBytes(localPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading local file '{localPath}': {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine("No Excel file available for conversion (download failed and local file not found).");
                return;
            }
        }

        // Load the workbook from the byte array
        try
        {
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                Workbook workbook = new Workbook(ms);
                // Convert and save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);
            }

            Console.WriteLine($"Conversion completed. PDF saved as {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
